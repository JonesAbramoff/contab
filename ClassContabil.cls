VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContabil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private WithEvents objContaButton As CommandButton
Attribute objContaButton.VB_VarHelpID = -1
Private WithEvents objCclButton As CommandButton
Attribute objCclButton.VB_VarHelpID = -1
Private WithEvents objHistButton As CommandButton
Attribute objHistButton.VB_VarHelpID = -1

Private WithEvents objEventoConta As AdmEvento
Attribute objEventoConta.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoHist As AdmEvento
Attribute objEventoHist.VB_VarHelpID = -1

'Variáveis pública para chaveamento de tela --> Jaques 13/1/2000
Public objTela As Object
Public objGridInt As AdmGrid
Public WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Public WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Public lNumIntDocOrigemExt As Long
Public lNumNotaFiscalExt As Long
Public iSubTipoTransacaoExt As Integer
Public iTrocouSubTipoTransacao As Integer

Private sSiglaModulo As String
Private sOrigemCusto As String 'origem dos lançamentos de custo
Private sNomeModulo As String
Private iGrid_Conta_Col As Integer
Private iGrid_Ccl_Col As Integer
Private iGrid_Debito_Col As Integer
Private iGrid_Credito_Col As Integer
Private iGrid_SeqContraPartida_Col As Integer
Private iGrid_Historico_Col As Integer
Private iGrid_Aglutina_Col As Integer
Private iGrid_Gerencial_Col As Integer
Private iContabilidade_Alterada As Integer
Private iExercicio As Integer
Private iPeriodo As Integer
Private iCodTransacao As Integer
Private iClienteFornecedor As Integer 'indica se se trata de um cliente ou um fornecedor associado ao lancamento
Private iGeraModelo As Integer
Private iObrigatoriedade As Integer
'Private lDoc As Long
'Private lDocOriginal As Long


Private Const GRID_COLUNA_CONTA = 1
Private Const GRID_COLUNA_CCL = 2

Public Function Contabil_Inicializa_Contabilidade3(objTela1 As Object, objGridInt1 As AdmGrid, objEventoLote1 As AdmEvento, objEventoDoc1 As AdmEvento, sSiglaModulo1 As String, Optional iSubTipoTransacao As Integer = 0) As Long
'usado para inicializar a contabilidade em telas pesadas cuja inicialização da contabilidade está sendo postergada

Dim lErro As Long

On Error GoTo Erro_Contabil_Inicializa_Contabilidade3

    iGeraModelo = 0

    Set objEventoConta = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoHist = New AdmEvento

    Set objEventoLote1 = New AdmEvento

    Set objEventoLote = objEventoLote1

    Set objEventoDoc1 = New AdmEvento

    Set objEventoDoc = objEventoDoc1

    Set objTela = objTela1
    
    iSubTipoTransacaoExt = iSubTipoTransacao

    Set objGridInt1 = New AdmGrid

    Set objGridInt = objGridInt1

    sSiglaModulo = sSiglaModulo1

    sNomeModulo = gcolModulo.Nome(sSiglaModulo)
    
    sOrigemCusto = gcolModulo.sOrigemCusto(sSiglaModulo)

    objTela.Controls("CTBOrigem").Caption = sNomeModulo

    'Tela em questão
    Set objGridInt.objForm = objTela1
    
    Contabil_Inicializa_Contabilidade3 = SUCESSO
    
    Exit Function
    
Erro_Contabil_Inicializa_Contabilidade3:

    Contabil_Inicializa_Contabilidade3 = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147802)

    End Select
    
    Exit Function

End Function

Public Function Contabil_Inicializa_Contabilidade2() As Long
'usado para inicializar a contabilidade em telas pesadas cuja inicialização da contabilidade está sendo postergada

Dim lErro As Long

    lErro = SUCESSO

    iGeraModelo = 0

    If objGridInt.objGrid Is Nothing Then
        lErro = Contabil_Inicializa_ContabilidadeAux()
    End If

    Contabil_Inicializa_Contabilidade2 = lErro

End Function

Public Function Contabil_Inicializa_Contabilidade(objTela1 As Object, objGridInt1 As AdmGrid, objEventoLote1 As AdmEvento, objEventoDoc1 As AdmEvento, sSiglaModulo1 As String, Optional iSubTipoTransacao As Integer = 0) As Long

Dim lErro As Long

On Error GoTo Erro_Contabil_Inicializa_Contabilidade

    iGeraModelo = 0

    Set objEventoConta = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoHist = New AdmEvento

    Set objEventoLote1 = New AdmEvento

    Set objEventoLote = objEventoLote1

    Set objEventoDoc1 = New AdmEvento

    Set objEventoDoc = objEventoDoc1

    Set objTela = objTela1
    
    iSubTipoTransacaoExt = iSubTipoTransacao

    Set objGridInt1 = New AdmGrid

    Set objGridInt = objGridInt1

    sSiglaModulo = sSiglaModulo1

    sNomeModulo = gcolModulo.Nome(sSiglaModulo)
    
    sOrigemCusto = gcolModulo.sOrigemCusto(sSiglaModulo)

    objTela.Controls("CTBOrigem").Caption = sNomeModulo

    'Tela em questão
    Set objGridInt.objForm = objTela1

    lErro = Contabil_Inicializa_ContabilidadeAux()
    If lErro <> SUCESSO Then Error 60761
    
    Contabil_Inicializa_Contabilidade = SUCESSO

    Exit Function

Erro_Contabil_Inicializa_Contabilidade:

    Contabil_Inicializa_Contabilidade = Err

    Select Case Err

        Case 60761

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147803)

    End Select

    Exit Function

End Function

Private Function Contabil_Inicializa_ContabilidadeAux() As Long

Dim colPadraoContab As New Collection
Dim objPadraoContab As ClassPadraoContab
Dim objTransacaoCTB As New ClassTransacaoCTB
Dim iCodTransacao1 As Integer
Dim lErro As Long

On Error GoTo Erro_Contabil_Inicializa_ContabilidadeAux

'    If giTipoVersao = VERSAO_LIGHT Then
'        lErro = Contabil_Inicializa_Versao_Light()
'        If lErro <> SUCESSO Then Error 60780
'    End If

    objTransacaoCTB.sSigla = sSiglaModulo
    objTransacaoCTB.sTransacao = objTela.Name
    objTransacaoCTB.iSubTipo = iSubTipoTransacaoExt
    
    lErro = CF("TransacaoCTB_Le", objTransacaoCTB)
    If lErro <> SUCESSO Then Error 36370

    If giContabGerencial = 1 Then
        objGridInt.iGerencial = objTransacaoCTB.iGerencial
    End If

    'inicializa o grid de contabilizacao
    lErro = Contabil_Inicializa_GridContabil()
    If lErro <> SUCESSO Then Error 14215

    objTela.Controls("CTBLabelContas").Visible = False
    objTela.Controls("CTBLabelCcl").Visible = False
    objTela.Controls("CTBTvwContas").Visible = False
    objTela.Controls("CTBTvwCcls").Visible = False
    objTela.Controls("CTBListHistoricos").Visible = False
    objTela.Controls("CTBLabelHistoricos").Visible = False
    objTela.Controls("CTBDocumento").Mask = "######"
    objTela.Controls("CTBDocumento").MaxLength = 6

    If objTela.Controls("CTBTvwContas").Nodes.Count = 0 Then

        If objContaButton Is Nothing Then
             Set objContaButton = objTela.Controls.Add("VB.CommandButton", "CTBContaBtn", objTela.Controls("CTBTvwContas").Container)
             With objContaButton
               .Visible = True
               .Width = 2000
               .Caption = "&Plano de Contas"
               .top = 3200
               .left = 6600
               .Height = 400
               .FontBold = True
            End With
        End If
        
'        'Inicializa a Arvore de Plano de Contas
'        lErro = CF("Carga_Arvore_Conta_Modulo", objTela.Controls("CTBTvwContas").Nodes, sSiglaModulo)
'        If lErro <> SUCESSO Then Error 36232
    
        If objContab.giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
    
'            'Inicializa a Lista de Centros de Custo
'            lErro = Contabil_Carga_Arvore_Ccl(objTela.Controls("CTBTvwCcls").Nodes)
'            If lErro <> SUCESSO Then Error 36233
    
            If objCclButton Is Nothing Then
                 Set objCclButton = objTela.Controls.Add("VB.CommandButton", "CTBCclBtn", objTela.Controls("CTBTvwContas").Container)
                 With objCclButton
                   .Visible = True
                   .Width = 2000
                   .Caption = "&Centros de Custo"
                   .top = 3700
                   .left = 6600
                   .Height = 400
                   .FontBold = True
                End With
            End If
        
        End If
    
'        'Inicializa a Lista de Historicos
'        lErro = Contabil_Carga_Lista_Historico()
'        If lErro <> SUCESSO Then Error 36234
    
        If objHistButton Is Nothing Then
             Set objHistButton = objTela.Controls.Add("VB.CommandButton", "CTBHistBtn", objTela.Controls("CTBTvwContas").Container)
             With objHistButton
               .Visible = True
               .Width = 2000
               .Caption = "&Históricos"
               .top = 4200
               .left = 6600
               .Height = 400
               .FontBold = True
            End With
        End If
    
        If objContab.giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
            objTela.Controls("CTBCclLabel").Visible = True
            objTela.Controls("CTBCclDescricao").Visible = True
        End If

    End If

    iObrigatoriedade = objTransacaoCTB.iObrigatoriedade
    iCodTransacao = objTransacaoCTB.iCodigo
    iClienteFornecedor = objTransacaoCTB.iCliFornecedor

    iCodTransacao1 = iCodTransacao

    'leitura dos modelos de contabilizacao referentes ao modulo/transacao em questao
    lErro = CF("PadraoContab_Le_Todos", sSiglaModulo, iCodTransacao1, colPadraoContab, iSubTipoTransacaoExt)
    If lErro <> SUCESSO Then Error 36226

    objTela.Controls("CTBModelo").Clear

    For Each objPadraoContab In colPadraoContab

        objTela.Controls("CTBModelo").AddItem objPadraoContab.sModelo
        objTela.Controls("CTBModelo").ItemData(objTela.Controls("CTBModelo").NewIndex) = objPadraoContab.iPadrao

    Next

    iContabilidade_Alterada = 0
    
    Contabil_Inicializa_ContabilidadeAux = SUCESSO
    
    Exit Function

Erro_Contabil_Inicializa_ContabilidadeAux:

    Contabil_Inicializa_ContabilidadeAux = Err
    
    Select Case Err
    
        Case 14215, 36226, 36232, 36233, 36234, 36370, 60780

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147804)

    End Select
    
    Exit Function
    
End Function

Private Function Contabil_Inicializa_Versao_Light() As Long


On Error GoTo Erro_Contabil_Inicializa_Versao_Light

'    objTela.Controls("CTBListHistoricos").TabStop = False
'    objTela.Controls("CTBListHistoricos").Left = -20000
'    objTela.Controls("CTBTvwCcls").TabStop = False
'    objTela.Controls("CTBTvwCcls").Left = -20000
'    objTela.Controls("CTBTvwContas").TabStop = False
'    objTela.Controls("CTBTvwContas").Left = -20000
'    objTela.Controls("CTBLabelHistoricos").Left = -20000
'    objTela.Controls("CTBLabelContas").Left = -20000
'    objTela.Controls("CTBLabelCcl").Left = -20000
'    objTela.Controls("CTBBotaoLimparGrid").TabStop = False
'    objTela.Controls("CTBBotaoLimparGrid").Left = -20000
'    objTela.Controls("CTBBotaoModeloPadrao").TabStop = False
'    objTela.Controls("CTBBotaoModeloPadrao").Left = -20000
'    objTela.Controls("CTBModelo").TabStop = False
'    objTela.Controls("CTBModelo").Left = -20000
'    objTela.Controls("CTBLabel1").Left = -20000
    objTela.Controls("CTBLote").TabStop = False
    objTela.Controls("CTBLote").left = -20000
    objTela.Controls("CTBLabelLote").left = -20000
'    objTela.Controls("CTBLancAutomatico").TabStop = False
'    objTela.Controls("CTBLancAutomatico").Left = -20000
'    objTela.Controls("CTBBotaoImprimir").TabStop = False
'    objTela.Controls("CTBBotaoImprimir").Left = -20000
'    objTela.Controls("CTBUpDown").TabStop = False
'    objTela.Controls("CTBUpDown").Left = -20000
'    objTela.Controls("CTBDocumento").BackColor = vbButtonFace
'    objTela.Controls("CTBDocumento").Enabled = False
'    objTela.Controls("CTBDataContabil").BackColor = vbButtonFace
'    objTela.Controls("CTBDataContabil").Enabled = False
'    objTela.Controls("CTBLabel21").Left = 45
'    objTela.Controls("CTBOrigem").Left = 750
'    objTela.Controls("CTBLabel8").Left = 225
'    objTela.Controls("CTBDataContabil").Left = 750
'    objTela.Controls("CTBLabelDoc").Left = 2700
'    objTela.Controls("CTBDocumento").Left = 3810
'    objTela.Controls("CTBLabel13").Left = 2850
'    objTela.Controls("CTBExercicio").Left = 3795
'    objTela.Controls("CTBLabel14").Left = 5910
'    objTela.Controls("CTBPeriodo").Left = 6690
'    objTela.Controls("CTBHistorico").Width = 3090
'    objTela.Controls("CTBConta").Enabled = False
'    objTela.Controls("CTBCcl").Enabled = False
'    objTela.Controls("CTBCredito").Enabled = False
'    objTela.Controls("CTBDebito").Enabled = False
'    objTela.Controls("CTBSeqContraPartida").Enabled = False
'    objTela.Controls("CTBHistorico").Enabled = False
'    objTela.Controls("CTBAglutina").Enabled = False

    objTela.Controls("CTBListHistoricos").Clear
    objTela.Controls("CTBTvwContas").Nodes.Clear
    objTela.Controls("CTBModelo").Clear
    If objContab.giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        objTela.Controls("CTBTvwCcls").Nodes.Clear
    End If

    Contabil_Inicializa_Versao_Light = SUCESSO
    
    Exit Function
    
Erro_Contabil_Inicializa_Versao_Light:

    Contabil_Inicializa_Versao_Light = Err

    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147805)

    End Select
    
    Exit Function

End Function

Private Function Contabil_Inicializa_GridContabil(Optional ByVal bConsulta As Boolean = False) As Long

Dim lErro As Long, obj2 As Object, sTipoControle As String
Dim iIndice As Integer

On Error GoTo Erro_Contabil_Inicializa_GridContabil


    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Conta")
    If objContab.giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colColuna.Add ("CCusto")
    objGridInt.colColuna.Add ("Débito")
    objGridInt.colColuna.Add ("Crédito")
    objGridInt.colColuna.Add ("C.P.")
    objGridInt.colColuna.Add ("Histórico")
    objGridInt.colColuna.Add ("Aglutina")
    If giContabGerencial = 1 Then objGridInt.colColuna.Add ("Status")

   'campos de edição do grid
    objGridInt.colCampo.Add (objTela.Controls("CTBConta").Name)
    If objContab.giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colCampo.Add (objTela.Controls("CTBCcl").Name)
    objGridInt.colCampo.Add (objTela.Controls("CTBDebito").Name)
    objGridInt.colCampo.Add (objTela.Controls("CTBCredito").Name)
    objGridInt.colCampo.Add (objTela.Controls("CTBSeqContraPartida").Name)
    objGridInt.colCampo.Add (objTela.Controls("CTBHistorico").Name)
    objGridInt.colCampo.Add (objTela.Controls("CTBAglutina").Name)
    If giContabGerencial = 1 Then objGridInt.colCampo.Add (objTela.Controls("CTBGerencial").Name)

    'indica onde estão situadas as colunas do grid
    If objContab.giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        iGrid_Conta_Col = 1
        iGrid_Ccl_Col = 2
        iGrid_Debito_Col = 3
        iGrid_Credito_Col = 4
        iGrid_SeqContraPartida_Col = 5
        iGrid_Historico_Col = 6
        iGrid_Aglutina_Col = 7
    Else
        iGrid_Conta_Col = 1
        '999 indica que não está sendo usado
        iGrid_Ccl_Col = 999
        iGrid_Debito_Col = 2
        iGrid_Credito_Col = 3
        iGrid_SeqContraPartida_Col = 4
        iGrid_Historico_Col = 5
        iGrid_Aglutina_Col = 6
        objTela.Controls("CTBCcl").Visible = False
    End If

    If giContabGerencial = 1 Then
        iGrid_Gerencial_Col = iGrid_Aglutina_Col + 1
    Else
        objTela.Controls("CTBGerencial").Visible = False
    End If
    
    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 14216

    objGridInt.objGrid = objTela.Controls("CTBGridContabil")

    'todas as linhas do grid
    objGridInt.objGrid.Rows = MAX_LANCAMENTOS_POR_DOC + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5

    objTela.Controls("CTBGridContabil").ColWidth(0) = 400

    If Not bConsulta Then
    
        objGridInt.objGrid.Width = 9000
        objTela.Controls("CTBHistorico").Width = 2950
        objTela.Controls("CTBConta").Width = 1500
        objTela.Controls("CTBCcl").Width = 1000
    
        objGridInt.objGrid.top = objGridInt.objGrid.top + 120
        objTela.Controls("CTBLancAutomatico").top = objTela.Controls("CTBLancAutomatico").top + 120
        For Each obj2 In objTela.Controls
            sTipoControle = TypeName(obj2)
            If sTipoControle = "Label" Then
                If obj2.Caption = "Lançamentos" Then obj2.top = obj2.top + 120
            End If
        Next
    
    End If
    
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGridInt)

    'Posiciona os painéis totalizadores
    objTela.Controls("CTBTotalDebito").top = objGridInt.objGrid.top + objGridInt.objGrid.Height
    objTela.Controls("CTBTotalDebito").left = objGridInt.objGrid.left
    For iIndice = 0 To iGrid_Debito_Col - 1
        objTela.Controls("CTBTotalDebito").left = objTela.Controls("CTBTotalDebito").left + objGridInt.objGrid.ColWidth(iIndice) + objGridInt.objGrid.GridLineWidth + 20
    Next

    objTela.Controls("CTBTotalDebito").Width = objGridInt.objGrid.ColWidth(iGrid_Debito_Col)

    objTela.Controls("CTBTotalCredito").top = objTela.Controls("CTBTotalDebito").top
    objTela.Controls("CTBTotalCredito").left = objTela.Controls("CTBTotalDebito").left + objTela.Controls("CTBTotalDebito").Width + objGridInt.objGrid.GridLineWidth
    objTela.Controls("CTBTotalCredito").Width = objGridInt.objGrid.ColWidth(iGrid_Credito_Col)

    'teste p/evitar que o label de totais fique por baixo do grid
    If objTela.Controls("CTBTotalDebito").Height > objTela.Controls("CTBLabelTotais").Height Then
        objTela.Controls("CTBLabelTotais").top = objTela.Controls("CTBTotalDebito").top + (objTela.Controls("CTBTotalDebito").Height - objTela.Controls("CTBLabelTotais").Height) / 2
    Else
        objTela.Controls("CTBLabelTotais").top = objTela.Controls("CTBTotalDebito").top
    End If
    
    objTela.Controls("CTBLabelTotais").left = objTela.Controls("CTBTotalDebito").left - objTela.Controls("CTBLabelTotais").Width

    objTela.Controls("CTBTotalCredito").Alignment = vbRightJustify
    objTela.Controls("CTBTotalDebito").Alignment = vbRightJustify

    Contabil_Inicializa_GridContabil = SUCESSO

    Exit Function

Erro_Contabil_Inicializa_GridContabil:

    Contabil_Inicializa_GridContabil = Err

    Select Case Err

        Case 14216

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147806)

    End Select

    Exit Function

End Function

Function Contabil_BotaoImprimir_Click() As Long

Dim lErro As Long, objRelTela As New ClassRelTela, iIndice1 As Integer
Dim colTemp As Collection, objLancamento As New ClassLancamento_Detalhe
Dim colLancamento_Detalhe As New Collection
Dim sOrigem As String, sExercicio As String, sPeriodo As String
Dim lDoc As Long, dtData As Date, lLote As Long
Dim sNomeTsk As String
Dim objUsoCcl As New ClassContab

On Error GoTo Erro_Contabil_BotaoImprimir_Click
    
    lErro = objRelTela.Iniciar("REL_LANC_CTB")
    If lErro <> SUCESSO Then Error 44411
  
    If objUsoCcl.giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        sNomeTsk = "Lanc_CCL"
    Else
        sNomeTsk = "Lanc_CTB"
    End If
    
    'obter dados comuns a todas as linhas do grid
    sOrigem = objContab.gobjColOrigem.Origem(objTela.Controls("CTBOrigem").Caption)
    sExercicio = objTela.Controls("CTBExercicio").Caption
    sPeriodo = objTela.Controls("CTBPeriodo").Caption
    
    lDoc = StrParaLong(objTela.Controls("CTBDocumento").ClipText)
    dtData = StrParaDate(objTela.Controls("CTBDataContabil").Text)
    lLote = StrParaLong(objTela.Controls("CTBLote").ClipText)
    
    'Preenche Objeto Lançamento_Detalhe
    lErro = Contabil_Grid_Lancamento_Detalhe(colLancamento_Detalhe, 0, 0, 0, -1, "")
    If lErro <> SUCESSO Then Error 44412
    
    For iIndice1 = 1 To colLancamento_Detalhe.Count
    
        Set objLancamento = colLancamento_Detalhe.Item(iIndice1)
        
        Set colTemp = New Collection
        
        'incluir os valores na mesma ordem da tabela RelTelaCampos no dicdados
        
        Call colTemp.Add(sOrigem)
        Call colTemp.Add(sExercicio)
        Call colTemp.Add(sPeriodo)
        Call colTemp.Add(lDoc)
        Call colTemp.Add(iIndice1)
        Call colTemp.Add(lLote)
        Call colTemp.Add(dtData)
        Call colTemp.Add(objLancamento.sConta)
        Call colTemp.Add(objLancamento.sCcl)
        Call colTemp.Add(objLancamento.sHistorico)
        Call colTemp.Add(objLancamento.dValor)
        Call colTemp.Add(objLancamento.iSeqContraPartida)

        lErro = objRelTela.IncluirRegistro(colTemp)
        If lErro <> SUCESSO Then Error 44413
    
    Next
    
    lErro = objRelTela.ExecutarRel(sNomeTsk)
    If lErro <> SUCESSO Then Error 44414
    
    Contabil_BotaoImprimir_Click = SUCESSO
    
    Exit Function
    
Erro_Contabil_BotaoImprimir_Click:

    Contabil_BotaoImprimir_Click = Err
    
    Select Case Err
          
        Case 44411, 44412, 44413, 44414
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147807)
     
    End Select
    
    Exit Function
    
End Function

Function Contabil_BotaoModeloPadrao_Click() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iAchou As Integer
Dim colLancamentos As New ClassColLancamentos

On Error GoTo Erro_Contabil_BotaoModeloPadrao_Click

    'pesquisa na combobox Modelo o modelo padrão
    For iIndice = 0 To objTela.Controls("CTBModelo").ListCount - 1
        If objTela.Controls("CTBModelo").ItemData(iIndice) = PADRAOCONTAB_MODELO_PADRAO Then
            iAchou = 1
            Exit For
        End If
    Next

    'se não encontrou o modelo padrão ==> mensagem de erro
    If iAchou <> 1 Then Error 36227

    objTela.Controls("CTBModelo").ListIndex = -1
    objTela.Controls("CTBModelo").ListIndex = iIndice

    Contabil_BotaoModeloPadrao_Click = SUCESSO

    Exit Function

Erro_Contabil_BotaoModeloPadrao_Click:

    Contabil_BotaoModeloPadrao_Click = Err

    Select Case Err

        Case 36227
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODELO_CONTAB_SEM_PADRAO", Err, sSiglaModulo, objTela.Name)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147808)

    End Select

    Exit Function

End Function

Function Contabil_Modelo_Click() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iAchou As Integer
Dim colLancamentos As New ClassColLancamentos

On Error GoTo Erro_Contabil_Modelo_Click

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        If objTela.Controls("CTBModelo").ListIndex >= 0 And iGeraModelo = 0 Then
    
            iContabilidade_Alterada = CONTABILIDADE_ALTERADA
    
            'retorna os lançamentos do Modulo/Transacao passados como parametro. Os lançamentos são retornados em colLancamentos
            lErro = CF("Gera_Lancamentos", sSiglaModulo, objTela, objTela.Controls("CTBModelo").Text, colLancamentos, iSubTipoTransacaoExt)
            If lErro <> SUCESSO Then Error 36228
    
            lErro = Contabil_GridContabil_Preenche(colLancamentos)
            If lErro <> SUCESSO Then Error 36230
            
            objTela.Controls("CTBLancAutomatico").Value = MARCADO
    
        End If

    End If

    Contabil_Modelo_Click = SUCESSO

    Exit Function

Erro_Contabil_Modelo_Click:

    Contabil_Modelo_Click = Err

    Select Case Err

        Case 36228, 36230

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147809)

    End Select

    Exit Function

End Function

Function Contabil_GridContabil_Preenche(colLancamentos As ClassColLancamentos) As Long
'Preenche o grid contábil com os dados passados por colLancamentos

Dim lErro As Long
Dim iIndice As Integer
Dim objLancamento As ClassLancamentos
Dim dColunaSoma As Double

On Error GoTo Erro_Contabil_GridContabil_Preenche

    Call Contabil_Limpa_GridContabil1

    'Número de Parcelas
    objGridInt.iLinhasExistentes = colLancamentos.Count

    'Coloca os valores do lancamento no Grid
    For iIndice = 1 To colLancamentos.Count

        Set objLancamento = colLancamentos.Item(iIndice)

        If Len(objLancamento.sConta) > 0 Then
            objTela.Controls("CTBConta").PromptInclude = False
            objTela.Controls("CTBConta").Text = objLancamento.sConta
            objTela.Controls("CTBConta").PromptInclude = True
            objGridInt.objGrid.TextMatrix(iIndice, iGrid_Conta_Col) = objTela.Controls("CTBConta").Text
        End If

        If objContab.giSetupUsoCcl = CCL_USA_EXTRACONTABIL And Len(objLancamento.sCcl) > 0 Then
            objTela.Controls("CTBCcl").PromptInclude = False
            objTela.Controls("CTBCcl").Text = objLancamento.sCcl
            objTela.Controls("CTBCcl").PromptInclude = True
            objGridInt.objGrid.TextMatrix(iIndice, iGrid_Ccl_Col) = objTela.Controls("CTBCcl").Text
        End If

        If objLancamento.dCredito <> 0 Then objGridInt.objGrid.TextMatrix(iIndice, iGrid_Credito_Col) = Format(objLancamento.dCredito, "Standard")

        If objLancamento.dDebito <> 0 Then objGridInt.objGrid.TextMatrix(iIndice, iGrid_Debito_Col) = Format(objLancamento.dDebito, "Standard")

        If objLancamento.iSeqContraPartida <> 0 Then objGridInt.objGrid.TextMatrix(iIndice, iGrid_SeqContraPartida_Col) = CStr(objLancamento.iSeqContraPartida)
        
        objGridInt.objGrid.TextMatrix(iIndice, iGrid_Historico_Col) = objLancamento.sHistorico

        objGridInt.objGrid.TextMatrix(iIndice, iGrid_Aglutina_Col) = objLancamento.iAglutina

        If giContabGerencial = 1 Then objGridInt.objGrid.TextMatrix(iIndice, iGrid_Gerencial_Col) = objLancamento.iGerencial

    Next

    dColunaSoma = Contabil_GridColuna_Soma(iGrid_Debito_Col)
    objTela.Controls("CTBTotalDebito") = Format(dColunaSoma, "Standard")
    dColunaSoma = Contabil_GridColuna_Soma(iGrid_Credito_Col)
    objTela.Controls("CTBTotalCredito") = Format(dColunaSoma, "Standard")

    Call Grid_Refresh_Checkbox(objGridInt)

    Contabil_GridContabil_Preenche = SUCESSO

    Exit Function

Erro_Contabil_GridContabil_Preenche:

    Contabil_GridContabil_Preenche = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147810)

    End Select

End Function

Private Function Contabil_Carga_Arvore_Ccl(colNodes As Nodes) As Long
'move os dados de centro de custo/lucro do banco de dados para a arvore colNodes.

Dim objNode As Node
Dim colCcl As New Collection
Dim objCcl As ClassCcl
Dim lErro As Long
Dim sCclMascarado As String
Dim sCcl As String
Dim sCclPai As String

On Error GoTo Erro_Contabil_Carga_Arvore_Ccl

    lErro = CF("Ccl_Le_Todos", colCcl)
    If lErro <> SUCESSO Then Error 36235

    For Each objCcl In colCcl

        sCclMascarado = String(STRING_CCL, 0)

        lErro = objContab.Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 36236

        If objCcl.iTipoCcl = CCL_ANALITICA Then
            sCcl = "A" & objCcl.sCcl
        Else
            sCcl = "S" & objCcl.sCcl
        End If

        sCclPai = String(STRING_CCL, 0)

        'retorna o centro de custo/lucro "pai" do centro de custo/lucro em questão, se houver
        lErro = objContab.Mascara_RetornaCclPai(objCcl.sCcl, sCclPai)
        If lErro <> SUCESSO Then Error 36237

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

    Contabil_Carga_Arvore_Ccl = SUCESSO

    Exit Function

Erro_Contabil_Carga_Arvore_Ccl:

    Contabil_Carga_Arvore_Ccl = Err

    Select Case Err

        Case 36235

        Case 36236
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)

        Case 36237
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclPai", Err, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147811)

    End Select

    Exit Function

End Function

Private Function Contabil_Carga_Lista_Historico() As Long
'move os dados de historico padrão do banco de dados para a arvore colNodes.

Dim colHistPadrao As New Collection
Dim objHistPadrao As ClassHistPadrao
Dim lErro As Long

On Error GoTo Erro_Contabil_Carga_Lista_Historico

    lErro = CF("HistPadrao_Le_Todos", colHistPadrao)
    If lErro <> SUCESSO Then Error 36241

    For Each objHistPadrao In colHistPadrao

        objTela.Controls("CTBListHistoricos").AddItem CStr(objHistPadrao.iHistPadrao) & SEPARADOR & objHistPadrao.sDescHistPadrao
        objTela.Controls("CTBListHistoricos").ItemData(objTela.Controls("CTBListHistoricos").NewIndex) = objHistPadrao.iHistPadrao

    Next

    Contabil_Carga_Lista_Historico = SUCESSO

    Exit Function

Erro_Contabil_Carga_Lista_Historico:

    Contabil_Carga_Lista_Historico = Err

    Select Case Err

        Case 36241

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147812)

    End Select

    Exit Function

End Function

Public Function Contabil_Saida_Celula(Optional objGridIntParam As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Contabil_Saida_Celula

    If Not (objGridInt Is Nothing) Then

        If Not (objGridIntParam Is Nothing) Then
    
            'verifica se o grid passado como parametro é o mesmo ao grid ao qual este objeto se refere. Se não for ==> retorna sucesso.
            If Not (objGridIntParam.objGrid Is objGridInt.objGrid) Then
                Contabil_Saida_Celula = SUCESSO
                Exit Function
            End If

        End If

        If objGridInt.objGrid = objTela.Controls("CTBGridContabil") Then
    
            Select Case objTela.Controls("CTBGridContabil").col
    
                Case iGrid_Conta_Col
    
                    lErro = Contabil_Saida_Celula_Conta()
                    If lErro <> SUCESSO Then gError 36243
    
                Case iGrid_Ccl_Col
    
                    lErro = Contabil_Saida_Celula_Ccl()
                    If lErro <> SUCESSO Then gError 36244
    
                Case iGrid_Credito_Col
    
                    lErro = Contabil_Saida_Celula_Credito()
                    If lErro <> SUCESSO Then gError 36245
    
                Case iGrid_Debito_Col
    
                    lErro = Contabil_Saida_Celula_Debito()
                    If lErro <> SUCESSO Then gError 36246
    
                Case iGrid_SeqContraPartida_Col
    
                    lErro = Contabil_Saida_Celula_SeqContraPartida()
                    If lErro <> SUCESSO Then gError 44051
    
                Case iGrid_Historico_Col
    
                    lErro = Contabil_Saida_Celula_Historico()
                    If lErro <> SUCESSO Then gError 36247
    
                Case iGrid_Aglutina_Col
    
                    lErro = Contabil_Saida_Celula_Aglutina()
                    If lErro <> SUCESSO Then gError 36276
    
                Case iGrid_Gerencial_Col
    
                    lErro = Contabil_Saida_Celula_Gerencial()
                    If lErro <> SUCESSO Then gError 188058
    
            End Select
    
        End If

    End If
    
    Contabil_Saida_Celula = SUCESSO

    Exit Function

Erro_Contabil_Saida_Celula:

    Contabil_Saida_Celula = gErr

    Select Case gErr

        Case 36243, 36244, 36245, 36246, 36247, 36276, 44051, 188058

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147813)

    End Select

    Exit Function

End Function

Private Function Contabil_Saida_Celula_Conta() As Long
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
Dim iIndice As Integer
Dim lPosicaoSeparador As Long

On Error GoTo Erro_Contabil_Saida_Celula_Conta

    Set objGridInt.objControle = objTela.Controls("CTBConta")

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", objTela.Controls("CTBConta").Text, objTela.Controls("CTBConta").ClipText, objPlanoConta, sSiglaModulo)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 44039
    
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = objContab.Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 20642
        
        objTela.Controls("CTBConta").PromptInclude = False
        objTela.Controls("CTBConta").Text = sContaMascarada
        objTela.Controls("CTBConta").PromptInclude = True
    
        'Se a Conta possui um Histórico Padrão "default" coloca na tela
        If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Historico_Col))) = 0 And objPlanoConta.iHistPadrao <> 0 Then
                        
            For iIndice = 0 To objTela.Controls("CTBListHistoricos").ListCount - 1
                If objTela.Controls("CTBListHistoricos").ItemData(iIndice) = objPlanoConta.iHistPadrao Then
                    objTela.Controls("CTBListHistoricos").ListIndex = iIndice
                    lPosicaoSeparador = InStr(objTela.Controls("CTBListHistoricos").Text, SEPARADOR)
                    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Historico_Col) = Mid(objTela.Controls("CTBListHistoricos").Text, lPosicaoSeparador + 1)
                    Exit For
                End If
            Next
            
        End If
    
    
    'se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", objTela.Controls("CTBConta").Text, sContaFormatada, objPlanoConta, sSiglaModulo)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 36248

        'conta não cadastrada
        If lErro = 5700 Then Error 36249

        'Se a Conta possui um Histórico Padrão "default" coloca na tela
        If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Historico_Col))) = 0 And objPlanoConta.iHistPadrao <> 0 Then
                        
            For iIndice = 0 To objTela.Controls("CTBListHistoricos").ListCount - 1
                If objTela.Controls("CTBListHistoricos").ItemData(iIndice) = objPlanoConta.iHistPadrao Then
                    objTela.Controls("CTBListHistoricos").ListIndex = iIndice
                    lPosicaoSeparador = InStr(objTela.Controls("CTBListHistoricos").Text, SEPARADOR)
                    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Historico_Col) = Mid(objTela.Controls("CTBListHistoricos").Text, lPosicaoSeparador + 1)
                    Exit For
                End If
            Next
            
        End If

    End If

    'se a conta foi preenchida
    If Len(objTela.Controls("CTBConta").ClipText) > 0 Then

        'se utiliza centro de custo extra-contabil
        If objContab.giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then

            'se o centro de custo foi preenchido
            If Len(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Ccl_Col)) > 0 Then

                'verifica se a associação da conta com o centro de custo está cadastrado
                sCcl = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Ccl_Col)

                lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
                If lErro <> SUCESSO Then Error 36250

                objContaCcl.sConta = sContaFormatada
                objContaCcl.sCcl = sCclFormatada

                lErro = CF("ContaCcl_Le", objContaCcl)
                If lErro <> SUCESSO And lErro <> 5871 Then Error 36251

                'associação Conta x Centro de Custo/Lucro não cadastrada
                If lErro = 5871 Then Error 36252

            End If

        End If

        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

        objTela.Controls("CTBContaDescricao").Caption = objPlanoConta.sDescConta

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 36253

    Contabil_Saida_Celula_Conta = SUCESSO

    Exit Function

Erro_Contabil_Saida_Celula_Conta:

    Contabil_Saida_Celula_Conta = Err

    Select Case Err

        Case 36248, 36250, 36251, 36253, 44039
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 36249
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", objTela.Controls("CTBConta").Text)

            If vbMsgRes = vbYes Then

                objPlanoConta.sConta = sContaFormatada

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela1(objGridInt)

                Call Chama_Tela("PlanoConta", objPlanoConta)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 36252
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", objTela.Controls("CTBConta").Text, sCcl)

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela1(objGridInt)

                Call Chama_Tela("ContaCcl", objContaCcl)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147814)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Contabil_Saida_Celula_Ccl() As Long
'faz a critica da celula ccl do grid que está deixando de ser a corrente

Dim sCclFormatada As String
Dim sContaFormatada As String
Dim lErro As Long
Dim iContaPreenchida As Integer
Dim objContaCcl As New ClassContaCcl
Dim sConta As String
Dim vbMsgRes As VbMsgBoxResult
Dim objCcl As New ClassCcl

On Error GoTo Erro_Contabil_Saida_Celula_Ccl

    Set objGridInt.objControle = objTela.Controls("CTBCcl")

    'critica o formato do ccl, sua presença no BD e capacidade de receber lançamentos
    lErro = CF("Ccl_Critica", objTela.Controls("CTBCcl").Text, sCclFormatada, objCcl)
    If lErro <> SUCESSO And lErro <> 5703 Then Error 36254

    'se o centro de custo/lucro não estiver cadastrado
    If lErro = 5703 Then Error 36255

    'se o centro de custo foi preenchido
    If Len(objTela.Controls("CTBCcl").ClipText) > 0 Then

        'se a conta foi informada
        If Len(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Conta_Col)) > 0 Then

            'verificar se a associação da conta com o centro de custo em questão está cadastrada
            sConta = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Conta_Col)

            lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then Error 36256

            objContaCcl.sConta = sContaFormatada
            objContaCcl.sCcl = sCclFormatada

            lErro = CF("ContaCcl_Le", objContaCcl)
            If lErro <> SUCESSO And lErro <> 5871 Then Error 36257

            'associação Conta x Centro de Custo/Lucro não cadastrada
            If lErro = 5871 Then Error 36258

        End If

        objTela.Controls("CTBCclDescricao").Caption = objCcl.sDescCcl

        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 36259

    Contabil_Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Contabil_Saida_Celula_Ccl:

    Contabil_Saida_Celula_Ccl = Err

    Select Case Err

        Case 36254, 36256, 36257, 36259
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 36255
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", objTela.Controls("CTBCcl").Text)

            If vbMsgRes = vbYes Then

                objCcl.sCcl = sCclFormatada

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela1(objGridInt)

                Call Chama_Tela("CclTela", objCcl)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 36258
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", sConta, objTela.Controls("CTBCcl").Text)

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela1(objGridInt)

                Call Chama_Tela("ContaCcl", objContaCcl)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147815)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Contabil_Saida_Celula_Credito() As Long
'faz a critica da celula credito do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double

On Error GoTo Erro_Contabil_Saida_Celula_Credito

    Set objGridInt.objControle = objTela.Controls("CTBCredito")

    If Len(objTela.Controls("CTBCredito").Text) > 0 Then
        lErro = Valor_Critica(objTela.Controls("CTBCredito").Text)
        If lErro <> SUCESSO Then Error 36260
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 36261

    If Len(objTela.Controls("CTBCredito").Text) > 0 Then
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Debito_Col) = ""

        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    dColunaSoma = Contabil_GridColuna_Soma(iGrid_Debito_Col)
    objTela.Controls("CTBTotalDebito") = Format(dColunaSoma, "Standard")
    dColunaSoma = Contabil_GridColuna_Soma(iGrid_Credito_Col)
    objTela.Controls("CTBTotalCredito") = Format(dColunaSoma, "Standard")

    Contabil_Saida_Celula_Credito = SUCESSO

    Exit Function

Erro_Contabil_Saida_Celula_Credito:

    Contabil_Saida_Celula_Credito = Err

    Select Case Err

        Case 36260, 36261
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147816)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Contabil_Saida_Celula_Debito() As Long
'faz a critica da celula debito do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double

On Error GoTo Erro_Contabil_Saida_Celula_Debito

    Set objGridInt.objControle = objTela.Controls("CTBDebito")

    If Len(objTela.Controls("CTBDebito").Text) > 0 Then
        lErro = Valor_Critica(objTela.Controls("CTBDebito").Text)
        If lErro <> SUCESSO Then Error 36262
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 36263

    If Len(objTela.Controls("CTBDebito").Text) > 0 Then
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Credito_Col) = ""

        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    dColunaSoma = Contabil_GridColuna_Soma(iGrid_Credito_Col)
    objTela.Controls("CTBTotalCredito") = Format(dColunaSoma, "Standard")
    dColunaSoma = Contabil_GridColuna_Soma(iGrid_Debito_Col)
    objTela.Controls("CTBTotalDebito") = Format(dColunaSoma, "Standard")

    Contabil_Saida_Celula_Debito = SUCESSO

    Exit Function

Erro_Contabil_Saida_Celula_Debito:

    Contabil_Saida_Celula_Debito = Err

    Select Case Err

        Case 36262, 36263
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147817)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Contabil_Saida_Celula_SeqContraPartida() As Long
'faz a critica da celula sequencial de contra partida do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Contabil_Saida_Celula_SeqContraPartida

    Set objGridInt.objControle = objTela.Controls("CTBSeqContraPartida")
    
    If Len(objTela.Controls("CTBSeqContraPartida").Text) > 0 Then
    
        If objGridInt.objGrid.Row = CInt(objTela.Controls("CTBSeqContraPartida").Text) Then Error 44052
        
        If CInt(objTela.Controls("CTBSeqContraPartida").Text) > objGridInt.iLinhasExistentes Then Error 44053
        
        If CInt(objTela.Controls("CTBSeqContraPartida").Text) <= 0 Then Error 44054
    
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44056
                
    Contabil_Saida_Celula_SeqContraPartida = SUCESSO
    
    Exit Function
    
Erro_Contabil_Saida_Celula_SeqContraPartida:

    Contabil_Saida_Celula_SeqContraPartida = Err
    
    Select Case Err
    
        Case 44052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTRAPARTIDA_NAO_MESMO_LANCAMENTO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 44053, 44054
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTRAPARTIDA_LANCAMENTO_INEXISTENTE", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 44056
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147818)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
    End Select

    Exit Function

End Function

Private Function Contabil_Saida_Celula_Historico() As Long
'faz a critica da celula historico do grid que está deixando de ser a corrente

Dim sValor As String
Dim lErro As Long
Dim objHistPadrao As ClassHistPadrao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Contabil_Saida_Celula_Historico

    Set objHistPadrao = New ClassHistPadrao

    Set objGridInt.objControle = objTela.Controls("CTBHistorico")

    If left(objTela.Controls("CTBHistorico").Text, 1) = CARACTER_HISTPADRAO Then

        sValor = Trim(Mid(objTela.Controls("CTBHistorico").Text, 2))

        lErro = Valor_Inteiro_Critica(sValor)
        If lErro <> SUCESSO Then Error 36264

        objHistPadrao.iHistPadrao = CInt(sValor)

        lErro = CF("HistPadrao_Le", objHistPadrao)
        If lErro <> SUCESSO And lErro <> 5446 Then Error 36265

        If lErro = 5446 Then Error 36266

        objTela.Controls("CTBHistorico").Text = objHistPadrao.sDescHistPadrao

    End If

    If Len(objTela.Controls("CTBHistorico").Text) > 0 Then
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 36267

    Contabil_Saida_Celula_Historico = SUCESSO

    Exit Function

Erro_Contabil_Saida_Celula_Historico:

    Contabil_Saida_Celula_Historico = Err

    Select Case Err

        Case 36264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", Err, sValor)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 36265, 36267
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 36266
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_HISTPADRAO_INEXISTENTE", objHistPadrao.iHistPadrao)

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela1(objGridInt)

                Call Chama_Tela("HistoricoPadrao", objHistPadrao)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147819)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Contabil_Saida_Celula_Aglutina() As Long
'faz a critica da celula(checkbox) aglutina do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Contabil_Saida_Celula_Aglutina

    Set objGridInt.objControle = objTela.Controls("CTBAglutina")

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 36277

    Contabil_Saida_Celula_Aglutina = SUCESSO

    Exit Function

Erro_Contabil_Saida_Celula_Aglutina:

    Contabil_Saida_Celula_Aglutina = Err

    Select Case Err

        Case 36277

            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147820)

    End Select

    Exit Function

End Function

Private Function Contabil_Saida_Celula_Gerencial() As Long
'faz a critica da celula(checkbox) aglutina do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Contabil_Saida_Celula_Gerencial

    Set objGridInt.objControle = objTela.Controls("CTBGerencial")

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188059

    Contabil_Saida_Celula_Gerencial = SUCESSO

    Exit Function

Erro_Contabil_Saida_Celula_Gerencial:

    Contabil_Saida_Celula_Gerencial = gErr

    Select Case gErr

        Case 188059
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 188060)

    End Select

    Exit Function

End Function

Function Contabil_GridColuna_Soma(iColuna As Integer) As Double
'soma o conteudo da coluna passada como parametro e retorna o total

Dim dAcumulador As Double
Dim iLinha As Integer

    dAcumulador = 0

    For iLinha = 1 To objGridInt.iLinhasExistentes
        If Len(objGridInt.objGrid.TextMatrix(iLinha, iColuna)) > 0 Then
            dAcumulador = dAcumulador + CDbl(objGridInt.objGrid.TextMatrix(iLinha, iColuna))
        End If
    Next

    Contabil_GridColuna_Soma = dAcumulador

End Function

Function Contabil_Conta_Change() As Long

    objTela.iAlterado = REGISTRO_ALTERADO
    iContabilidade_Alterada = CONTABILIDADE_ALTERADA
'mario
    objTela.Controls("CTBLancAutomatico").Value = DESMARCADO
    Contabil_Conta_Change = SUCESSO

End Function

Function Contabil_Conta_GotFocus() As Long
'trata o evento gotfocus do campo conta

Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_Contabil_Conta_GotFocus

    
    Call Grid_Campo_Recebe_Foco(objGridInt)



'    objTela.Controls("CTBTvwContas").Visible = True
'    objTela.Controls("CTBTvwCcls").Visible = False
    objTela.Controls("CTBListHistoricos").Visible = False
'    objTela.Controls("CTBLabelContas").Visible = True
'    objTela.Controls("CTBLabelCcl").Visible = False
    objTela.Controls("CTBLabelHistoricos").Visible = False

    sConta = objTela.Controls("CTBGridContabil").TextMatrix(objTela.Controls("CTBGridContabil").Row, objTela.Controls("CTBGridContabil").col)

    If Len(sConta) > 0 Then

        lErro = Contabil_Conta_Exibe_Descricao(sConta)
        If lErro <> SUCESSO Then Error 14213

    Else

        objTela.Controls("CTBContaDescricao") = ""

    End If
    
    Contabil_Conta_GotFocus = SUCESSO

    Exit Function

Erro_Contabil_Conta_GotFocus:

    Contabil_Conta_GotFocus = Err

    Select Case Err

        Case 14213

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147821)

    End Select

    Exit Function

End Function

Function Contabil_Conta_KeyPress(iKeyAscii As Integer) As Long
'trata o evento keypress do campo conta

    Call Grid_Trata_Tecla_Campo_Contabil(iKeyAscii, objGridInt)

    Contabil_Conta_KeyPress = SUCESSO

End Function

Function Contabil_Conta_Validate(Cancel As Boolean) As Long
'trata o evento validate do campo conta

Dim lErro As Long

    Set objGridInt.objControle = objTela.Controls("CTBConta")
    
    lErro = Grid_Campo_Libera_Foco_Contabil(objGridInt)
    If lErro <> SUCESSO Then Cancel = True

    Contabil_Conta_Validate = SUCESSO

End Function

Function Contabil_Ccl_Change() As Long

    objTela.iAlterado = REGISTRO_ALTERADO
    iContabilidade_Alterada = CONTABILIDADE_ALTERADA
'mario
    objTela.Controls("CTBLancAutomatico").Value = DESMARCADO
    Contabil_Ccl_Change = SUCESSO

End Function

Function Contabil_Ccl_GotFocus() As Long
'trata o evento gotfocus do campo ccl

Dim sCcl As String
Dim lErro As Long

On Error GoTo Erro_Contabil_Ccl_GotFocus

    Call Grid_Campo_Recebe_Foco(objGridInt)

'    objTela.Controls("CTBTvwCcls").Visible = True
'    objTela.Controls("CTBTvwContas").Visible = False
'    objTela.Controls("CTBListHistoricos").Visible = False
'    objTela.Controls("CTBLabelContas").Visible = False
'    objTela.Controls("CTBLabelCcl").Visible = True
'    objTela.Controls("CTBLabelHistoricos").Visible = False

    sCcl = objTela.Controls("CTBGridContabil").TextMatrix(objTela.Controls("CTBGridContabil").Row, objTela.Controls("CTBGridContabil").col)

    'Coloca descricao de Ccl no panel
    If Len(sCcl) > 0 Then

        lErro = Contabil_Ccl_Exibe_Descricao(sCcl)
        If lErro <> SUCESSO Then Error 14214

    Else

        objTela.Controls("CTBCclDescricao") = ""

    End If

    Contabil_Ccl_GotFocus = SUCESSO

    Exit Function

Erro_Contabil_Ccl_GotFocus:

    Contabil_Ccl_GotFocus = Err

    Select Case Err

        Case 14214

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147822)

    End Select

    Exit Function

End Function

Function Contabil_Ccl_KeyPress(iKeyAscii As Integer) As Long
'trata o evento keypress associado ao campo ccl

    Call Grid_Trata_Tecla_Campo_Contabil(iKeyAscii, objGridInt)

    Contabil_Ccl_KeyPress = SUCESSO

End Function

Function Contabil_Ccl_Validate(Cancel As Boolean) As Long
'trata o evento validate associado ao campo ccl

Dim lErro As Long

    Set objGridInt.objControle = objTela.Controls("CTBCcl")
    lErro = Grid_Campo_Libera_Foco_Contabil(objGridInt)
    If lErro <> SUCESSO Then Cancel = True

    Contabil_Ccl_Validate = SUCESSO

End Function

Function Contabil_Credito_Change() As Long

    objTela.iAlterado = REGISTRO_ALTERADO
    iContabilidade_Alterada = CONTABILIDADE_ALTERADA
'mario
   objTela.Controls("CTBLancAutomatico").Value = DESMARCADO
    Contabil_Credito_Change = SUCESSO

End Function

Function Contabil_Credito_GotFocus() As Long
'trata o evento gotfocus associado ao campo credito

    Call Grid_Campo_Recebe_Foco(objGridInt)

    Contabil_Credito_GotFocus = SUCESSO

End Function

Function Contabil_Credito_KeyPress(iKeyAscii As Integer) As Long
'trata o evento keypress associado ao campo credito

    Call Grid_Trata_Tecla_Campo_Contabil(iKeyAscii, objGridInt)

    Contabil_Credito_KeyPress = SUCESSO

End Function

Function Contabil_Credito_Validate(Cancel As Boolean) As Long
'trata o evento validate associado ao campo credito

Dim lErro As Long

    Set objGridInt.objControle = objTela.Controls("CTBCredito")
    lErro = Grid_Campo_Libera_Foco_Contabil(objGridInt)
    If lErro <> SUCESSO Then Cancel = True

    Contabil_Credito_Validate = SUCESSO

End Function

Function Contabil_Debito_Change() As Long

    objTela.iAlterado = REGISTRO_ALTERADO
    iContabilidade_Alterada = CONTABILIDADE_ALTERADA
'mario
   objTela.Controls("CTBLancAutomatico").Value = DESMARCADO
    Contabil_Debito_Change = SUCESSO

End Function

Function Contabil_Debito_GotFocus() As Long
'trata o evento gotfocus associado ao campo debito

    Call Grid_Campo_Recebe_Foco(objGridInt)

    Contabil_Debito_GotFocus = SUCESSO

End Function

Function Contabil_Debito_KeyPress(iKeyAscii As Integer) As Long
'trata o evento keypress associado ao campo debito

    Call Grid_Trata_Tecla_Campo_Contabil(iKeyAscii, objGridInt)

    Contabil_Debito_KeyPress = SUCESSO

End Function

Function Contabil_Debito_Validate(Cancel As Boolean) As Long
'trata o evento validate associado ao campo debito

Dim lErro As Long

    Set objGridInt.objControle = objTela.Controls("CTBDebito")
    lErro = Grid_Campo_Libera_Foco_Contabil(objGridInt)
    If lErro <> SUCESSO Then Cancel = True

    Contabil_Debito_Validate = SUCESSO

End Function

Function Contabil_SeqContraPartida_Change() As Long

    objTela.iAlterado = REGISTRO_ALTERADO
    iContabilidade_Alterada = CONTABILIDADE_ALTERADA
    Contabil_SeqContraPartida_Change = SUCESSO

End Function

Function Contabil_SeqContraPartida_GotFocus() As Long
'trata o evento gotfocus associado ao campo sequencial de contra partida

    Call Grid_Campo_Recebe_Foco(objGridInt)

    Contabil_SeqContraPartida_GotFocus = SUCESSO

End Function

Function Contabil_SeqContraPartida_KeyPress(iKeyAscii As Integer) As Long
'trata o evento keypress associado ao campo sequencial de contra partida

    Call Grid_Trata_Tecla_Campo_Contabil(iKeyAscii, objGridInt)

    Contabil_SeqContraPartida_KeyPress = SUCESSO

End Function

Function Contabil_SeqContraPartida_Validate(Cancel As Boolean) As Long
'trata o evento validate associado ao campo sequencial de contra partida

Dim lErro As Long

    Set objGridInt.objControle = objTela.Controls("CTBSeqContraPartida")
    lErro = Grid_Campo_Libera_Foco_Contabil(objGridInt)
    If lErro <> SUCESSO Then Cancel = True

    Contabil_SeqContraPartida_Validate = SUCESSO

End Function

Function Contabil_Historico_Change() As Long

    objTela.iAlterado = REGISTRO_ALTERADO
    iContabilidade_Alterada = CONTABILIDADE_ALTERADA
'mario
    objTela.Controls("CTBLancAutomatico").Value = DESMARCADO
    Contabil_Historico_Change = SUCESSO

End Function

Function Contabil_Historico_GotFocus() As Long
'trata o evento gotfocus associado ao campo historico

Dim iPos As Integer

    Call Grid_Campo_Recebe_Foco(objGridInt)

'    objTela.Controls("CTBListHistoricos").Visible = True
'    objTela.Controls("CTBTvwCcls").Visible = False
'    objTela.Controls("CTBTvwContas").Visible = False
'    objTela.Controls("CTBLabelContas").Visible = False
'    objTela.Controls("CTBLabelCcl").Visible = False
'    objTela.Controls("CTBLabelHistoricos").Visible = True

    If Len(objTela.Controls("CTBHistorico").Text) > 0 Then
        iPos = InStr(objTela.Controls("CTBHistorico").Text, CARACTER_HISTORICO_PARAM)
        If iPos > 0 Then
            objTela.Controls("CTBHistorico").SelStart = iPos - 1
            objTela.Controls("CTBHistorico").SelLength = 1
        End If
    End If

    Contabil_Historico_GotFocus = SUCESSO

End Function

Function Contabil_Historico_KeyPress(iKeyAscii As Integer) As Long
'trata o evento keypress associado ao campo historico

Dim iInicio As Integer
Dim iPos As Integer
Dim sValor As String
Dim lErro As Long
Dim objHistPadrao As New ClassHistPadrao

On Error GoTo Erro_Contabil_Historico_KeyPress

    'se digitou ENTER
    If iKeyAscii = vbKeyReturn Then
        
        If Len(objTela.Controls("CTBHistorico").Text) > 0 Then
        
            If left(objTela.Controls("CTBHistorico").Text, 1) = CARACTER_HISTPADRAO Then
            
                sValor = Trim(Mid(objTela.Controls("CTBHistorico").Text, 2))
                
                lErro = Valor_Inteiro_Critica(sValor)
                If lErro <> SUCESSO Then Error 44079
                
                objHistPadrao.iHistPadrao = CInt(sValor)
                        
                lErro = CF("HistPadrao_Le", objHistPadrao)
                If lErro <> SUCESSO And lErro <> 5446 Then Error 44080
                
                If lErro = 5446 Then Error 44081
        
                objTela.Controls("CTBHistorico").Text = objHistPadrao.sDescHistPadrao
                objTela.Controls("CTBHistorico").SelStart = 0
                
            End If
    
            If objTela.Controls("CTBHistorico").SelText = CARACTER_HISTORICO_PARAM Then
                iInicio = objTela.Controls("CTBHistorico").SelStart + 2
            Else
                iInicio = objTela.Controls("CTBHistorico").SelStart
            End If
        
            If iInicio = 0 Then iInicio = 1
        
            iPos = InStr(iInicio, objTela.Controls("CTBHistorico").Text, CARACTER_HISTORICO_PARAM)
            If iPos > 0 Then
                objTela.Controls("CTBHistorico").SelStart = iPos - 1
                objTela.Controls("CTBHistorico").SelLength = 1
                Exit Function
            End If
        End If
    End If

    Call Grid_Trata_Tecla_Campo_Contabil(iKeyAscii, objGridInt)

    Contabil_Historico_KeyPress = SUCESSO

    Exit Function
    
Erro_Contabil_Historico_KeyPress:

    Contabil_Historico_KeyPress = Err
    
    Select Case Err
    
        Case 44079
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", Err, sValor)
        
        Case 44080

        Case 44081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_CADASTRADO", Err, objHistPadrao.iHistPadrao)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147823)
    
    End Select

    Exit Function

End Function

Function Contabil_Historico_Validate(Cancel As Boolean) As Long
'trata o evento validate associado ao campo historico

Dim lErro As Long

    Set objGridInt.objControle = objTela.Controls("CTBHistorico")
    lErro = Grid_Campo_Libera_Foco_Contabil(objGridInt)
    If lErro <> SUCESSO Then Cancel = True

    Contabil_Historico_Validate = SUCESSO

End Function

Function Contabil_Aglutina_Click() As Long

    objTela.iAlterado = REGISTRO_ALTERADO
    iContabilidade_Alterada = CONTABILIDADE_ALTERADA
'mario
   objTela.Controls("CTBLancAutomatico").Value = DESMARCADO
    Contabil_Aglutina_Click = SUCESSO

End Function

Function Contabil_Aglutina_GotFocus() As Long
'trata o evento gotfocus associado ao campo aglutina

    Call Grid_Campo_Recebe_Foco(objGridInt)

    Contabil_Aglutina_GotFocus = SUCESSO

End Function

Function Contabil_Aglutina_KeyPress(iKeyAscii As Integer) As Long
'trata o evento keypress associado ao campo aglutina

    Call Grid_Trata_Tecla_Campo_Contabil(iKeyAscii, objGridInt)

    Contabil_Aglutina_KeyPress = SUCESSO

End Function

Function Contabil_Aglutina_Validate(Cancel As Boolean) As Long
'trata o evento validate associado ao campo aglutina

Dim lErro As Long

    Set objGridInt.objControle = objTela.Controls("CTBAglutina")
    lErro = Grid_Campo_Libera_Foco_Contabil(objGridInt)
    If lErro <> SUCESSO Then Cancel = True
    
    Contabil_Aglutina_Validate = SUCESSO

End Function

Function Contabil_TvwContas_NodeClick(ByVal Node As MSComctlLib.Node) As Long
'trata o evento nodeclick associado a arvore tvwcontas

Dim sConta As String
Dim sCaracterInicial As String
Dim lPosicaoSeparador As Long
Dim lErro As Long
Dim sContaEnxuta As String
Dim objPlanoConta As New ClassPlanoConta
Dim sContaFormatada As String
Dim iIndice As Integer

On Error GoTo Erro_Contabil_TvwContas_NodeClick

    If objTela.Controls("CTBGridContabil").col = iGrid_Conta_Col Then

        sCaracterInicial = left(Node.Key, 1)

        If sCaracterInicial = "A" Then

            sConta = right(Node.Key, Len(Node.Key) - 1)

            sContaEnxuta = String(STRING_CONTA, 0)

            lErro = objContab.Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
            If lErro <> SUCESSO Then Error 36268

            objTela.Controls("CTBConta").PromptInclude = False
            objTela.Controls("CTBConta").Text = sContaEnxuta
            objTela.Controls("CTBConta").PromptInclude = True

            objTela.Controls("CTBGridContabil").TextMatrix(objTela.Controls("CTBGridContabil").Row, objTela.Controls("CTBGridContabil").col) = objTela.Controls("CTBConta").Text

            If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            End If

            'Preenche a Descricao da Conta
            lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
            objTela.Controls("CTBContaDescricao").Caption = Mid(Node.Text, lPosicaoSeparador + 1)

            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", objTela.Controls("CTBConta").Text, sContaFormatada, objPlanoConta, sSiglaModulo)
            If lErro <> SUCESSO And lErro <> 5700 Then Error 60758
                    
            'Conta não cadastrada
            If lErro = 5700 Then Error 60759

            'Se a Conta possui um Histórico Padrão "default" coloca na tela
            If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Historico_Col))) = 0 And objPlanoConta.iHistPadrao <> 0 Then
                            
                For iIndice = 0 To objTela.Controls("CTBListHistoricos").ListCount - 1
                    If objTela.Controls("CTBListHistoricos").ItemData(iIndice) = objPlanoConta.iHistPadrao Then
                        objTela.Controls("CTBListHistoricos").ListIndex = iIndice
                        lPosicaoSeparador = InStr(objTela.Controls("CTBListHistoricos").Text, SEPARADOR)
                        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Historico_Col) = Mid(objTela.Controls("CTBListHistoricos").Text, lPosicaoSeparador + 1)
                        Exit For
                    End If
                Next
                
            End If

        End If

    End If

    Contabil_TvwContas_NodeClick = SUCESSO

    Exit Function

Erro_Contabil_TvwContas_NodeClick:

    Contabil_TvwContas_NodeClick = Err

    Select Case Err

        Case 36268
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
            
        Case 60758
        
        Case 60759
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, objTela.Controls("CTBConta").Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147824)

    End Select

    Exit Function

End Function

Function Contabil_TvwContas_Expand(ByVal objNode As MSComctlLib.Node, colNodes As Nodes) As Long

Dim lErro As Long

On Error GoTo Erro_Contabil_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta_Modulo1", objNode, colNodes, sSiglaModulo)
        If lErro <> SUCESSO Then Error 44022
        
    End If
    
    Contabil_TvwContas_Expand = SUCESSO
    
    Exit Function
    
Erro_Contabil_TvwContas_Expand:

    Select Case Err
    
        Case 44022
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147825)
        
    End Select
        
    Exit Function
    
End Function


Function Contabil_TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node) As Long
'trata o evento nodeclick da arvore tvwccls

Dim sCcl As String
Dim sCclEnxuta As String
Dim lErro As Long
Dim lPosicaoSeparador As Long
Dim sCaracterInicial As String

On Error GoTo Erro_Contabil_TvwCcls_NodeClick

    If objTela.Controls("CTBGridContabil").col = iGrid_Ccl_Col Then

        sCaracterInicial = left(Node.Key, 1)

        If sCaracterInicial = "A" Then

            sCcl = right(Node.Key, Len(Node.Key) - 1)

            sCclEnxuta = String(STRING_CCL, 0)

            'volta mascarado apenas os caracteres preenchidos
            lErro = objContab.Mascara_RetornaCclEnxuta(sCcl, sCclEnxuta)
            If lErro <> SUCESSO Then Error 36275

            objTela.Controls("CTBCcl").PromptInclude = False
            objTela.Controls("CTBCcl").Text = sCclEnxuta
            objTela.Controls("CTBCcl").PromptInclude = True

            objTela.Controls("CTBGridContabil").TextMatrix(objTela.Controls("CTBGridContabil").Row, objTela.Controls("CTBGridContabil").col) = objTela.Controls("CTBCcl").Text

            If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            End If

            'Preenche a Descricao do centro de custo/lucro
            lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
            objTela.Controls("CTBCclDescricao").Caption = Mid(Node.Text, lPosicaoSeparador + 1)

        End If

    End If

    Contabil_TvwCcls_NodeClick = SUCESSO

    Exit Function

Erro_Contabil_TvwCcls_NodeClick:

    Contabil_TvwCcls_NodeClick = Err

    Select Case Err

        Case 36275
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", Err, sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147826)

    End Select

    Exit Function

End Function

Function Contabil_Conta_Exibe_Descricao(sConta As String) As Long
'exibe a descrição da conta no campo ContaDescricao. A conta passada como parametro deve estar mascarada

Dim sContaFormatada As String
Dim lErro As Long
Dim iContaPreenchida As Integer
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Contabil_Conta_Exibe_Descricao

    'Retorna conta formatada como no BD
    lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
    If lErro <> SUCESSO Then Error 36269

    If iContaPreenchida = CONTA_PREENCHIDA Then
    
        lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, 1)
        If lErro <> SUCESSO And lErro <> 6030 Then Error 36270
    
        If lErro = 6030 Then Error 36271
    
        objTela.Controls("CTBContaDescricao").Caption = objPlanoConta.sDescConta

    Else
    
        objTela.Controls("CTBContaDescricao").Caption = ""
    
    End If
    
    Contabil_Conta_Exibe_Descricao = SUCESSO

    Exit Function

Erro_Contabil_Conta_Exibe_Descricao:

    Contabil_Conta_Exibe_Descricao = Err

    Select Case Err

        Case 36269, 36270, 36271
            objTela.Controls("CTBContaDescricao") = ""

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147827)

    End Select

    Exit Function

End Function

Function Contabil_Ccl_Exibe_Descricao(sCcl As String) As Long
'exibe a descrição do centro de custo/lucro no campo CclDescricao. O ccl passado como parametro deve estar mascarado

Dim sCclFormatada As String
Dim sCclArvore As String
Dim objNode As Node
Dim lErro As Long
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl

On Error GoTo Erro_Contabil_Ccl_Exibe_Descricao

    'Retorna Ccl formatada como no BD
    lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then Error 36272

    If iCclPreenchida = CCL_PREENCHIDA Then
    
        objCcl.sCcl = sCclFormatada
    
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then Error 36273
    
        If lErro = 5599 Then Error 36274
    
        objTela.Controls("CTBCclDescricao").Caption = objCcl.sDescCcl

    Else
        objTela.Controls("CTBCclDescricao").Caption = ""
    End If

    Contabil_Ccl_Exibe_Descricao = SUCESSO

    Exit Function

Erro_Contabil_Ccl_Exibe_Descricao:

    Contabil_Ccl_Exibe_Descricao = Err

    Select Case Err

        Case 36272, 36273, 36274
            objTela.Controls("CTBCclDescricao") = ""

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147828)

    End Select

    Exit Function

End Function

Function Contabil_ListHistoricos_DblClick() As Long
'trata um duplo click na listbox de historicos

Dim lPosicaoSeparador As Long

    If objTela.Controls("CTBGridContabil").col = iGrid_Historico_Col Then

        lPosicaoSeparador = InStr(objTela.Controls("CTBListHistoricos").Text, SEPARADOR)
        objTela.Controls("CTBGridContabil").TextMatrix(objTela.Controls("CTBGridContabil").Row, objTela.Controls("CTBGridContabil").col) = Mid(objTela.Controls("CTBListHistoricos").Text, lPosicaoSeparador + 1)
        objTela.Controls("CTBHistorico").Text = Mid(objTela.Controls("CTBListHistoricos").Text, lPosicaoSeparador + 1)

        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    Contabil_ListHistoricos_DblClick = SUCESSO

End Function

Function Contabil_Lote_Change() As Long

    objTela.iAlterado = REGISTRO_ALTERADO
    iContabilidade_Alterada = CONTABILIDADE_ALTERADA
    Contabil_Lote_Change = SUCESSO

End Function

Function Contabil_Lote_GotFocus() As Long

Dim iAlterado As Integer

    iAlterado = objTela.iAlterado

    Call MaskEdBox_TrataGotFocus(objTela.Controls("CTBLote"), iAlterado)

    objTela.iAlterado = iAlterado

    Contabil_Lote_GotFocus = SUCESSO

End Function

Function Contabil_Lote_Validate(Cancel As Boolean, Parent As Object) As Long
'tratamento do validate do campo Lote

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objLote As New ClassLote
Dim sNomeExterno As String
Dim objExercicio As New ClassExercicio
Dim iLoteAtualizado As Integer
Dim colSelecao As Collection

On Error GoTo Erro_Contabil_Lote_Validate

    If Len(objTela.Controls("CTBLote").Text) > 0 And Len(objTela.Controls("CTBDataContabil").ClipText) > 0 Then

        objLote.iLote = CInt(objTela.Controls("CTBLote").Text)
        objLote.iFilialEmpresa = giFilialEmpresa
        objLote.sOrigem = sSiglaModulo

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(objTela.Controls("CTBDataContabil").Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 36278

        objLote.iExercicio = objPeriodo.iExercicio
        objLote.iPeriodo = objPeriodo.iPeriodo

        lErro = CF("LotePendente_Le", objLote)
        If lErro <> SUCESSO And lErro <> 5435 Then Error 36281

        'Se o lote pendente não está cadastrado
        If lErro = 5435 Then

            'pesquisa se é um lote atualizado
            lErro = CF("Lote_Le", objLote)
            If lErro <> SUCESSO And lErro <> 9293 Then Error 36325

            'se o lote não estiver cadastrado
            If lErro = 9293 Then Error 36282

        End If

        If objContab.giSetupLotePorPeriodo <> LOTE_INICIALIZADO_POR_PERIODO And objPeriodo.iPeriodo <> objLote.iPeriodo Then Error 36283

    End If

    Contabil_Lote_Validate = SUCESSO

    Exit Function

Erro_Contabil_Lote_Validate:

    Cancel = True
    
    If Not (Parent Is GL_objMDIForm.ActiveForm) Then
        Parent.Show
        Parent.SetFocus
    End If

    Contabil_Lote_Validate = Err

    Select Case Err

        Case 36278, 36281, 36325

        Case 36282
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_INEXISTENTE", objLote.iFilialEmpresa, objLote.iLote, objTela.Controls("CTBOrigem").Caption, objPeriodo.iPeriodo, objPeriodo.iExercicio)
            If vbMsgRes = vbYes Then
                'Se respondeu que deseja criar LOTE
                Call Chama_Tela("LoteTela", objLote, sSiglaModulo)
            End If

        Case 36283
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODOS_DIFERENTES", Err, objPeriodo.iPeriodo, objLote.iPeriodo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147829)

    End Select

    Exit Function

End Function

Function Contabil_DataContabil_Change() As Long

    objTela.iAlterado = REGISTRO_ALTERADO
    iContabilidade_Alterada = CONTABILIDADE_ALTERADA
    Contabil_DataContabil_Change = SUCESSO

End Function

Function Contabil_DataContabil_GotFocus() As Long

Dim iAlterado As Integer

    iAlterado = objTela.iAlterado

    Call MaskEdBox_TrataGotFocus(objTela.Controls("CTBDataContabil"), iAlterado)

    objTela.iAlterado = iAlterado

    Contabil_DataContabil_GotFocus = SUCESSO

End Function

Function Contabil_DataContabil_Validate(Cancel As Boolean, Parent As Object) As Long
'Tratamento do validate do campo DataContabil

Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim lDoc As Long
Dim sNomeExterno As String
Dim objLote As New ClassLote
Dim vbMsgRes As VbMsgBoxResult
Dim iLoteAtualizado As Integer
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Contabil_DataContabil_Validate

    If Len(objTela.Controls("CTBDataContabil").ClipText) > 0 Then

        lErro = Data_Critica(objTela.Controls("CTBDataContabil").Text)
        If lErro <> SUCESSO Then Error 36313

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(objTela.Controls("CTBDataContabil").Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 36314

        'Verifica se Exercicio está fechado
        lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
        If lErro <> SUCESSO And lErro <> 10083 Then Error 36315

        'Exercicio não cadastrado
        If lErro = 10083 Then Error 36316

        If objExercicio.iStatus = EXERCICIO_FECHADO Then Error 36317

        objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
        objPeriodosFilial.iExercicio = objPeriodo.iExercicio
        objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo
        objPeriodosFilial.sOrigem = objContab.gobjColOrigem.Origem(objTela.Controls("CTBOrigem").Caption)

        lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
        If lErro <> SUCESSO Then Error 36318

        If giDesconsideraFechamentoPeriodo <> MARCADO Then
            If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 36319
        End If

        'checa se o lote pertence ao periodo em questão
        If Len(objTela.Controls("CTBLote").Text) > 0 Then

            objLote.iLote = CInt(objTela.Controls("CTBLote").Text)

            objLote.iFilialEmpresa = giFilialEmpresa
            objLote.sOrigem = sSiglaModulo
            objLote.iExercicio = objPeriodo.iExercicio
            objLote.iPeriodo = objPeriodo.iPeriodo

            lErro = CF("LotePendente_Le", objLote)
            If lErro <> SUCESSO And lErro <> 5435 Then Error 36320

            'Se o lote pendente não está cadastrado
            If lErro = 5435 Then

                'pesquisa se é um lote atualizado
                lErro = CF("Lote_Le", objLote)
                If lErro <> SUCESSO And lErro <> 9293 Then Error 36324

                'se o lote não estiver cadastrado
                If lErro = 9293 Then Error 36321

            End If

            If objContab.giSetupLotePorPeriodo <> LOTE_INICIALIZADO_POR_PERIODO And objPeriodo.iPeriodo <> objLote.iPeriodo Then Error 36322

        End If


'        If objTela.Controls("CTBPeriodo").Caption <> objPeriodo.sNomeExterno Or objTela.Controls("CTBExercicio").Caption <> objExercicio.sNomeExterno Then
'
'            lErro = X.Voucher_Automatico(giFilialEmpresa, objExercicio.iExercicio, objPeriodo.iPeriodo, sSiglaModulo, lDoc)
'            If lErro <> SUCESSO Then Error 36323
'
'            objTela.Controls("CTBDocumento").Text = CStr(lDoc)
'
'        End If

        'Preenche campo de periodo
        objTela.Controls("CTBPeriodo").Caption = objPeriodo.sNomeExterno

        objTela.Controls("CTBExercicio").Caption = objExercicio.sNomeExterno

    Else

        objTela.Controls("CTBPeriodo").Caption = ""

        objTela.Controls("CTBExercicio").Caption = ""

    End If

    Contabil_DataContabil_Validate = SUCESSO

    Exit Function

Erro_Contabil_DataContabil_Validate:

    Cancel = True
    
    If Not (Parent Is GL_objMDIForm.ActiveForm) Then
        Parent.Show
        Parent.SetFocus
    End If
    
    Contabil_DataContabil_Validate = Err

    Select Case Err

        Case 36313

        Case 36314, 36315, 36318, 36320, 36323, 36324

        Case 36316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case 36317
            'Não é possível fazer lançamentos em exercício fechado
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objPeriodo.iExercicio)

        Case 36319
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)

        Case 36321
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_INEXISTENTE", objLote.iFilialEmpresa, objLote.iLote, sNomeModulo, objLote.iPeriodo, objLote.iExercicio)

            If vbMsgRes = vbYes Then
                'Se respondeu que deseja criar LOTE
                Call Chama_Tela("LoteTela", objLote, sSiglaModulo)
            End If

        Case 36322
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODOS_DIFERENTES", Err, objPeriodo.iPeriodo, objLote.iPeriodo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147830)

    End Select

    Exit Function

End Function

Private Function Contabil_Traz_Cabecalho_Tela() As Long

Dim sDescricao As String
Dim iPeriodoDoc As Integer
Dim iExercicioDoc As Integer
Dim iIndice As Integer
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim lErro As Long

On Error GoTo Erro_Contabil_Traz_Cabecalho_Tela

    'Inicializa Data
    objTela.Controls("CTBDataContabil").Text = Format(gdtDataAtual, "dd/mm/yy")

    'Coloca o periodo relativo a data na tela
    lErro = CF("Periodo_Le", gdtDataAtual, objPeriodo)
    If lErro <> SUCESSO Then Error 36284
    
    iPeriodo = objPeriodo.iPeriodo
    iExercicio = objPeriodo.iExercicio

    objTela.Controls("CTBPeriodo").Caption = objPeriodo.sNomeExterno

    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 36285

    If lErro = 10083 Then Error 36286

    objTela.Controls("CTBExercicio").Caption = objExercicio.sNomeExterno

    Contabil_Traz_Cabecalho_Tela = SUCESSO

    Exit Function

Erro_Contabil_Traz_Cabecalho_Tela:

    Contabil_Traz_Cabecalho_Tela = Err

    Select Case Err

        Case 36284, 36285

        Case 36286
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147831)

    End Select

    Exit Function

End Function

Function Contabil_Traz_Doc_Tela(lNumIntDoc As Long) As Long
'traz os dados do voucher do banco de dados para a tela

Dim lErro As Long
Dim lErro1 As Long
Dim colLancamentos As New Collection
Dim objLanc As ClassLancamento_Detalhe
Dim iLinha As Integer
Dim sDescricao As String
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim iIndice As Integer
Dim dAcumulador As Double
Dim sModelo As String
Dim dtData As Date
Dim Z As Object
Dim tPlanoConta As typePlanoConta
Dim objPlanoConta As New ClassPlanoConta
Dim iTeste As Integer
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Contabil_Traz_Doc_Tela

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        lErro = Contabil_Inicializa_Contabilidade2()
        If lErro <> SUCESSO Then Error 60781
    
        Call Contabil_Limpa_Contabilidade
    
        iContabilidade_Alterada = CONTABILIDADE_ALTERADA
    
        iFilialEmpresa = giFilialEmpresa
    
        'le os lançamentos pendentes contidos no documento em questão
        lErro = CF("LanPendente_Le_Doc3", giFilialEmpresa, iCodTransacao, lNumIntDoc, colLancamentos, sSiglaModulo, sModelo, dtData)
        If lErro <> SUCESSO And lErro <> 92019 Then Error 36299
    
        If giContabGerencial = 1 And ((giFilialEmpresa > giFilialAuxiliar And giFilialAuxiliar > 0) Or (giFilialEmpresa < Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
    
            'le os lançamentos pendentes contidos no documento em questão
            lErro1 = CF("LanPendente_Le_Doc3", giFilialEmpresa - giFilialAuxiliar, iCodTransacao, lNumIntDoc, colLancamentos, sSiglaModulo, sModelo, dtData)
            If lErro1 <> SUCESSO And lErro1 <> 92019 Then Error 36299
    
        Else
            lErro1 = 92019
    
        End If
    
        iGeraModelo = 1
    
        For iIndice = 0 To objTela.Controls("CTBModelo").ListCount - 1
            If objTela.Controls("CTBModelo").List(iIndice) = sModelo Then
                objTela.Controls("CTBModelo").ListIndex = iIndice
                Exit For
            End If
        Next
    
        iGeraModelo = 0
    
        'Inicializa Data
        If dtData <> 0 Then objTela.Controls("CTBDataContabil").Text = Format(dtData, "dd/mm/yy")
    
        'se não encontrou o documento
        If lErro = 92019 And lErro1 = 92019 Then
    
            'le os lançamentos contidos no documento em questão
            lErro = CF("Lancamento_Le_Doc4", giFilialEmpresa, iCodTransacao, lNumIntDoc, colLancamentos, sSiglaModulo, sModelo, dtData)
            If lErro <> SUCESSO And lErro <> 92024 Then Error 36300
    
            If giContabGerencial = 1 And ((giFilialEmpresa > giFilialAuxiliar And giFilialAuxiliar > 0) Or (giFilialEmpresa < Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
    
                'le os lançamentos contidos no documento em questão
                lErro1 = CF("Lancamento_Le_Doc4", giFilialEmpresa - giFilialAuxiliar, iCodTransacao, lNumIntDoc, colLancamentos, sSiglaModulo, sModelo, dtData)
                If lErro1 <> SUCESSO And lErro1 <> 92024 Then Error 36300
    
            Else
                lErro1 = 92024
    
            End If
            
            iGeraModelo = 1
    
            For iIndice = 0 To objTela.Controls("CTBModelo").ListCount - 1
                If objTela.Controls("CTBModelo").List(iIndice) = sModelo Then
                    objTela.Controls("CTBModelo").ListIndex = iIndice
                    Exit For
                End If
            Next
    
            iGeraModelo = 0
    
            'Inicializa Data
            If dtData <> 0 Then objTela.Controls("CTBDataContabil").Text = Format(dtData, "dd/mm/yy")
    
            'se não encontrou o documento
            If lErro = 92024 And lErro1 = 92024 Then Error 36326
    
        End If
    
        If lErro = SUCESSO Or lErro1 = SUCESSO Then
    
            Set objLanc = colLancamentos.Item(1)
    
            If objLanc.iLote <> 0 Then objTela.Controls("CTBLote").Text = CStr(objLanc.iLote)
    
            objTela.Controls("CTBDocumento").Text = CStr(objLanc.lDoc)
    
            'Coloca o periodo relativo a data na tela
            lErro = CF("Periodo_Le", objLanc.dtData, objPeriodo)
            If lErro <> SUCESSO Then Error 36301
    
            objTela.Controls("CTBPeriodo").Caption = objPeriodo.sNomeExterno
            iPeriodo = objPeriodo.iPeriodo
            iExercicio = objPeriodo.iExercicio
    
            'Coloca o exercicio na tela
            lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
            If lErro <> SUCESSO And lErro <> 10083 Then Error 36302
    
            'Se o exercicio não está cadastrado
            If lErro = 10083 Then Error 36303
    
            objTela.Controls("CTBExercicio").Caption = objExercicio.sNomeExterno
    
            'move os dados para a tela
            lErro = Contabil_Traz_Doc_Tela1(colLancamentos)
            If lErro <> SUCESSO Then Error 36371
    
            dAcumulador = Contabil_GridColuna_Soma(iGrid_Credito_Col)
            objTela.Controls("CTBTotalCredito").Caption = Format(dAcumulador, "Standard")
        
            dAcumulador = Contabil_GridColuna_Soma(iGrid_Debito_Col)
            objTela.Controls("CTBTotalDebito").Caption = Format(dAcumulador, "Standard")
            
        End If

    End If

    Contabil_Traz_Doc_Tela = SUCESSO

    Exit Function

Erro_Contabil_Traz_Doc_Tela:

    Contabil_Traz_Doc_Tela = Err

    Select Case Err

        Case 36299, 36300, 36301, 36302, 36326, 36371, 60781

        Case 36303
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147832)

    End Select

    iGeraModelo = 0

    Exit Function

End Function

Private Function Contabil_Traz_Doc_Tela1(colLancamentos As Collection) As Long
'traz os lançamentos para o grid

Dim lErro As Long
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim objLanc As ClassLancamento_Detalhe

On Error GoTo Erro_Contabil_Traz_Doc_Tela1

    For Each objLanc In colLancamentos

        'se pegar um lançamento de custo ==> pula
        If objLanc.sOrigem = sSiglaModulo Then

            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
    
            lErro = objContab.Mascara_RetornaContaEnxuta(objLanc.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 36304
    
            objTela.Controls("CTBConta").PromptInclude = False
            objTela.Controls("CTBConta").Text = sContaMascarada
            objTela.Controls("CTBConta").PromptInclude = True
    
            'coloca a conta na tela
            objTela.Controls("CTBGridContabil").TextMatrix(objLanc.iSeq, iGrid_Conta_Col) = objTela.Controls("CTBConta").Text
    
            If objContab.giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
    
                'mascara o centro de custo
                sCclMascarado = String(STRING_CCL, 0)
    
                If objLanc.sCcl <> "" Then
    
                    lErro = objContab.Mascara_RetornaCclEnxuta(objLanc.sCcl, sCclMascarado)
                    If lErro <> SUCESSO Then Error 36305
            
                    objTela.Controls("CTBCcl").PromptInclude = False
                    objTela.Controls("CTBCcl").Text = sCclMascarado
                    objTela.Controls("CTBCcl").PromptInclude = True
                    
                    'coloca o centro de custo na tela
                    objTela.Controls("CTBGridContabil").TextMatrix(objLanc.iSeq, iGrid_Ccl_Col) = objTela.Controls("CTBCcl").Text
                
                End If
    
            End If
    
            'coloca o valor na tela
            If objLanc.dValor > 0 Then
                objTela.Controls("CTBGridContabil").TextMatrix(objLanc.iSeq, iGrid_Credito_Col) = Format(objLanc.dValor, "Standard")
            Else
                objTela.Controls("CTBGridContabil").TextMatrix(objLanc.iSeq, iGrid_Debito_Col) = Format(-objLanc.dValor, "Standard")
            End If
            
            'coloca o sequencial de contra partida na tela, se houver
            If objLanc.iSeqContraPartida <> 0 Then objTela.Controls("CTBGridContabil").TextMatrix(objLanc.iSeq, iGrid_SeqContraPartida_Col) = CStr(objLanc.iSeqContraPartida)
                
            'coloca o historico na tela
            objTela.Controls("CTBGridContabil").TextMatrix(objLanc.iSeq, iGrid_Historico_Col) = objLanc.sHistorico
    
            'coloca a indicação de aglutinação na tela
            objTela.Controls("CTBGridContabil").TextMatrix(objLanc.iSeq, iGrid_Aglutina_Col) = CStr(objLanc.iAglutina)
    
            'coloca a indicação de aglutinação na tela
            If giContabGerencial = 1 Then objTela.Controls("CTBGridContabil").TextMatrix(objLanc.iSeq, iGrid_Gerencial_Col) = CStr(objLanc.iGerencial)
    
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            
        End If

    Next

    Call Grid_Refresh_Checkbox(objGridInt)

    Contabil_Traz_Doc_Tela1 = SUCESSO

    Exit Function

Erro_Contabil_Traz_Doc_Tela1:

    Contabil_Traz_Doc_Tela1 = Err

    Select Case Err

        Case 36304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objLanc.sConta)

        Case 36305
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objLanc.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147833)

    End Select

    Exit Function

End Function

Function Contabil_Limpa_Contabilidade() As Long

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        If Not (objGridInt.objGrid Is Nothing) Then
    
            Call Grid_Limpa(objGridInt)
            objTela.Controls("CTBTotalDebito").Caption = ""
            objTela.Controls("CTBTotalCredito").Caption = ""
            objTela.Controls("CTBContaDescricao").Caption = ""
            objTela.Controls("CTBCclDescricao").Caption = ""
            objTela.Controls("CTBDocumento").Text = ""
            objTela.Controls("CTBLote").Text = ""
            objTela.Controls("CTBDataContabil").Text = "  /  /  "
            objTela.Controls("CTBModelo").ListIndex = -1
            objTela.Controls("CTBPeriodo").Caption = ""
            objTela.Controls("CTBExercicio").Caption = ""
    '        lDoc = 0
    '        lDocOriginal = 0
            iContabilidade_Alterada = 0
        
        End If
    
    End If
    
    Contabil_Limpa_Contabilidade = SUCESSO

End Function

Function Contabil_Limpa_GridContabil() As Long

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        objTela.Controls("CTBModelo").ListIndex = -1
        Call Contabil_Limpa_GridContabil1
    
    End If
    
    Contabil_Limpa_GridContabil = SUCESSO

End Function

Function Contabil_Limpa_GridContabil1() As Long

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        Call Grid_Limpa(objGridInt)
        objTela.Controls("CTBTotalDebito").Caption = ""
        objTela.Controls("CTBTotalCredito").Caption = ""
        objTela.Controls("CTBContaDescricao").Caption = ""
        objTela.Controls("CTBCclDescricao").Caption = ""
    
    End If
    
    Contabil_Limpa_GridContabil1 = SUCESSO

End Function

Function Contabil_LabelLote_Click() As Long
'trata o click no label de Lote

Dim objLote As New ClassLote
Dim dtData As Date
Dim lErro As Long
Dim objPeriodo As New ClassPeriodo
Dim colSelecao As New Collection

On Error GoTo Erro_Contabil_LabelLote_Click

    'Obtém Periodo e Exercicio correspondentes à data
    If Len(objTela.Controls("CTBDataContabil").ClipText) > 0 Then
        dtData = CDate(objTela.Controls("CTBDataContabil").Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 36311

    Else
        objPeriodo.iExercicio = 0
        objPeriodo.iPeriodo = 0
    End If

    If Len(objTela.Controls("CTBLote").Text) = 0 Then
        objLote.iLote = 0
    Else
        objLote.iLote = CInt(objTela.Controls("CTBLote").Text)
    End If

    objLote.sOrigem = sSiglaModulo
    objLote.iExercicio = objPeriodo.iExercicio
    objLote.iPeriodo = objPeriodo.iPeriodo

    colSelecao.Add objLote.sOrigem
    colSelecao.Add giFilialEmpresa
    colSelecao.Add 0

    Call Chama_Tela("LotePendenteLista", colSelecao, objLote, objEventoLote)

    Contabil_LabelLote_Click = SUCESSO

    Exit Function

Erro_Contabil_LabelLote_Click:

    Contabil_LabelLote_Click = Err

    Select Case Err

        Case 36311

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147834)

    End Select

    Exit Function

End Function

Function Contabil_objEventoLote_evSelecao(obj1 As Object) As Long
'traz o lote selecionado para a tela

Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objLote As ClassLote
Dim iIndice As Integer
Dim sDescricao As String

On Error GoTo Erro_Contabil_objEventoLote_evSelecao

    Set objLote = obj1

    'Se estiver com a data preenchida ==> verificar se a data está dentro do periodo do lote
    If Len(objTela.Controls("CTBDataContabil").ClipText) > 0 Then

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(objTela.Controls("CTBDataContabil").Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 36329

        'se o periodo/exercicio não corresponde ao periodo/exercicio do lote ==> troca a data
        If objPeriodo.iExercicio <> objLote.iExercicio Or objPeriodo.iPeriodo <> objLote.iPeriodo Then

            'move a data inicial do lote, exercicio e periodo para a tela
            lErro = Contabil_Move_Data_Tela(objLote)
            If lErro <> SUCESSO Then Error 36330

        End If

    Else

        'se não estiver com a data preenchida
        'move a data inicial do lote, exercicio e periodo para a tela
        lErro = Contabil_Move_Data_Tela(objLote)
        If lErro <> SUCESSO Then Error 36331

    End If

    objTela.Controls("CTBLote").Text = CStr(objLote.iLote)

    objTela.Show

    Contabil_objEventoLote_evSelecao = SUCESSO

    Exit Function

Erro_Contabil_objEventoLote_evSelecao:

    Contabil_objEventoLote_evSelecao = Err

    Select Case Err

        Case 36329, 36330, 36331

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147835)

    End Select

    Exit Function

End Function

Function Contabil_LabelDoc_Click() As Long
'trata o click no label de Doc

Dim objLancamento_Detalhe As New ClassLancamento_Detalhe
Dim dtData As Date
Dim lErro As Long
Dim objPeriodo As New ClassPeriodo
Dim colSelecao As New Collection

On Error GoTo Erro_Contabil_LabelDoc_Click

    If Len(objTela.Controls("CTBDataContabil").ClipText) > 0 Then

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(objTela.Controls("CTBDataContabil").Text)
    
        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 36472
        
    Else
        objPeriodo.iExercicio = 0
        objPeriodo.iPeriodo = 0
    End If
    
    If Len(objTela.Controls("CTBDocumento").Text) = 0 Then
        objLancamento_Detalhe.lDoc = 0
    Else
        objLancamento_Detalhe.lDoc = CLng(objTela.Controls("CTBDocumento").ClipText)
    End If
    
    objLancamento_Detalhe.iFilialEmpresa = giFilialEmpresa
    objLancamento_Detalhe.sOrigem = sSiglaModulo
    objLancamento_Detalhe.iExercicio = objPeriodo.iExercicio
    objLancamento_Detalhe.iPeriodoLan = objPeriodo.iPeriodo
    objLancamento_Detalhe.iPeriodoLote = objPeriodo.iPeriodo
    
    If Len(objTela.Controls("CTBLote").Text) = 0 Then
        objLancamento_Detalhe.iLote = 0
    Else
        objLancamento_Detalhe.iLote = CInt(objTela.Controls("CTBLote").Text)
    End If
    
    objLancamento_Detalhe.iSeq = 0
    
    Call Chama_Tela("LanPendenteLista", colSelecao, objLancamento_Detalhe, objEventoDoc)

    Exit Function

Erro_Contabil_LabelDoc_Click:
    
    Select Case Err
    
        Case 36472
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147836)
        
    End Select
    
    Exit Function

End Function

Function Contabil_objEventoDoc_evSelecao(obj1 As Object) As Long
'??? Mario, completar isto aqui, se achar necessario
Dim objLancamento_Detalhe As ClassLancamento_Detalhe

    Set objLancamento_Detalhe = obj1
    
    If Not (objLancamento_Detalhe Is Nothing) Then
        
        objTela.Controls("CTBDocumento").Text = CStr(objLancamento_Detalhe.lDoc)
        
    End If
    
    objTela.Show
    
    Contabil_objEventoDoc_evSelecao = SUCESSO

End Function

Private Function Contabil_Move_Data_Tela(objLote As ClassLote) As Long

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim objPeriodo As New ClassPeriodo
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Contabil_Move_Data_Tela

    lErro = CF("Periodo_Le_ExercicioPeriodo", objLote.iExercicio, objLote.iPeriodo, objPeriodo)
    If lErro <> SUCESSO Then Error 36331

    'Verifica se Exercicio está fechado
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 36332

    'se o exercicio não estiver cadastrado
    If lErro = 10083 Then Error 36333

    If objExercicio.iStatus = EXERCICIO_FECHADO Then Error 36334

    objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
    objPeriodosFilial.iExercicio = objPeriodo.iExercicio
    objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo

    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then Error 36335

    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 36336
    End If
    
    objTela.Controls("CTBDataContabil").Text = Format(objPeriodo.dtDataInicio, "dd/mm/yy")

    objTela.Controls("CTBPeriodo").Caption = objPeriodo.sNomeExterno

    objTela.Controls("CTBExercicio").Caption = objExercicio.sNomeExterno

    Contabil_Move_Data_Tela = SUCESSO

    Exit Function

Erro_Contabil_Move_Data_Tela:

    Contabil_Move_Data_Tela = Err

    Select Case Err

        Case 36331, 36332, 36335

        Case 36333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case 36334
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objPeriodo.iExercicio)

        Case 36336
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147837)

    End Select

    Exit Function

End Function

Function Contabil_Documento_GotFocus() As Long

Dim iAlterado As Integer

    iAlterado = objTela.iAlterado

    Call MaskEdBox_TrataGotFocus(objTela.Controls("CTBDocumento"), iAlterado)

    objTela.iAlterado = iAlterado

    Contabil_Documento_GotFocus = SUCESSO

End Function

Function Contabil_Documento_Change() As Long

    objTela.iAlterado = REGISTRO_ALTERADO
    iContabilidade_Alterada = CONTABILIDADE_ALTERADA
    Contabil_Documento_Change = SUCESSO

End Function

Function Contabil_Carga_Modelo_Padrao1() As Long
'usado por telas que postergam a inicializacao da contabilidade

Dim lErro As Long

On Error GoTo Erro_Contabil_Carga_Modelo_Padrao1

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        lErro = Contabil_Inicializa_Contabilidade2()
        If lErro <> SUCESSO Then Error 60762
        
        lErro = Contabil_Carga_Modelo_Padrao()
        If lErro <> SUCESSO Then Error 60763
    
    End If
    
    Contabil_Carga_Modelo_Padrao1 = SUCESSO
    
    Exit Function
    
Erro_Contabil_Carga_Modelo_Padrao1:

    Contabil_Carga_Modelo_Padrao1 = Err
    
    Select Case Err

        Case 60762, 60763

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147838)

    End Select

    Exit Function
    
End Function

Function Contabil_Carga_Modelo_Padrao() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iAchou As Integer
Dim colLancamentos As New ClassColLancamentos
Dim lDoc As Long

On Error GoTo Erro_Contabil_Carga_Modelo_Padrao

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        lErro = Contabil_Inicializa_Contabilidade2()
        If lErro <> SUCESSO Then gError 60784
    
    '    If iTrocouSubTipoTransacao = 1 Then
    '        lErro = Contabil_Carga_Modelos()
    '        If lErro <> SUCESSO Then gError 71779
    '    End If
    
        If iContabilidade_Alterada <> CONTABILIDADE_ALTERADA Then
    
            lErro = Contabil_Gera_Cabecalho_Automatico()
            If lErro <> SUCESSO Then gError 59406
            
    '        'mostra número do proximo voucher(documento) disponível
    '        lErro = X.Voucher_Automatico(giFilialEmpresa, iExercicio, iPeriodo, sSiglaModulo, lDoc)
    '        If lErro <> SUCESSO Then gError 36363
    '
    '        objTela.Controls("CTBDocumento").Text = CStr(lDoc)
    
            'pesquisa na combobox Modelo o modelo padrão
            For iIndice = 0 To objTela.Controls("CTBModelo").ListCount - 1
                If objTela.Controls("CTBModelo").ItemData(iIndice) = PADRAOCONTAB_MODELO_PADRAO Then
                    iAchou = 1
                    Exit For
                End If
            Next
    
            If iAchou = 1 Then objTela.Controls("CTBModelo").ListIndex = iIndice
    
        'se o flag de geração de lançamentos automáticos estiver ligado ==> regera os lançamentos automáticos
    'mario
        ElseIf objTela.Controls("CTBLancAutomatico").Value = MARCADO Then
    
            Call Contabil_Limpa_GridContabil1
    
            
    
    '        If giTipoVersao = VERSAO_LIGHT Then
    '
    '            If objTela.Controls("CTBModelo").ListIndex = -1 Then
    '
    '                'pesquisa na combobox Modelo o modelo padrão
    '                For iIndice = 0 To objTela.Controls("CTBModelo").ListCount - 1
    '                    If objTela.Controls("CTBModelo").ItemData(iIndice) = PADRAOCONTAB_MODELO_PADRAO Then
    '                        iAchou = 1
    '                        Exit For
    '                    End If
    '                Next
    '
    '                If iAchou = 1 Then objTela.Controls("CTBModelo").ListIndex = iIndice
    '            Else
    '                Call Contabil_Modelo_Click
    '            End If
    '
    '        ElseIf giTipoVersao = VERSAO_FULL Then
            Call Contabil_Modelo_Click
    '        End If
    
        End If

    End If

    Contabil_Carga_Modelo_Padrao = SUCESSO

    Exit Function

Erro_Contabil_Carga_Modelo_Padrao:

    Contabil_Carga_Modelo_Padrao = gErr

    Select Case gErr

        Case 36352, 36363, 36410, 59406, 60784, 71779

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147839)

    End Select

    Exit Function

End Function

Private Function Contabil_Carga_Modelo_Padrao_Int() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iAchou As Integer
Dim colLancamentos As New ClassColLancamentos
Dim lDoc As Long
Dim objPeriodo As New ClassPeriodo

On Error GoTo Erro_Contabil_Carga_Modelo_Padrao_Int

    If iContabilidade_Alterada <> CONTABILIDADE_ALTERADA Then

        lErro = Contabil_Gera_Cabecalho_Automatico()
        If lErro <> SUCESSO Then Error 59407
        
'        'mostra número do proximo voucher(documento) disponível
'        lErro = X.Voucher_Automatico_Trans(giFilialEmpresa, iExercicio, iPeriodo, sSiglaModulo, lDoc)
'        If lErro <> SUCESSO Then Error 36363
'
'        objTela.Controls("CTBDocumento").Text = CStr(lDoc)

        'pesquisa na combobox Modelo o modelo padrão
        For iIndice = 0 To objTela.Controls("CTBModelo").ListCount - 1
            If objTela.Controls("CTBModelo").ItemData(iIndice) = PADRAOCONTAB_MODELO_PADRAO Then
                iAchou = 1
                Exit For
            End If
        Next

        If iAchou = 1 Then objTela.Controls("CTBModelo").ListIndex = iIndice

    End If

    Contabil_Carga_Modelo_Padrao_Int = SUCESSO

    Exit Function

Erro_Contabil_Carga_Modelo_Padrao_Int:

    Contabil_Carga_Modelo_Padrao_Int = Err

    Select Case Err

        Case 36362, 36363, 36410, 59407

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147840)

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

    'Lê a mascara das contas
    lErro = objContab.MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then Error 36337

    objTela.Controls("CTBConta").Mask = sMascaraConta

    'Se usa centro de custo/lucro ==> inicializa mascara de centro de custo/lucro
    If objContab.giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then

        sMascaraCcl = String(STRING_CCL, 0)

        'Lê a mascara dos centros de custo/lucro
        lErro = objContab.MascaraCcl(sMascaraCcl)
        If lErro <> SUCESSO Then Error 36338

        objTela.Controls("CTBCcl").Mask = sMascaraCcl

    End If

    Inicializa_Mascaras = SUCESSO

    Exit Function

Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err

    Select Case Err

        Case 36337, 36338

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147841)

    End Select

    Exit Function


End Function

Sub Contabil_GridContabil_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridInt, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridInt, objTela.iAlterado)
    End If

End Sub

Sub Contabil_GridContabil_EnterCell()

    Call Grid_Entrada_Celula(objGridInt, objTela.iAlterado)

End Sub

Sub Contabil_GridContabil_GotFocus()

    Call Grid_Recebe_Foco(objGridInt)

End Sub

Sub Contabil_GridContabil_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridInt, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridInt, objTela.iAlterado)
    End If

End Sub

Sub Contabil_GridContabil_KeyDown(KeyCode As Integer)

Dim dColunaSoma As Double
Dim lErro As Long

On Error GoTo Erro_Contabil_GridContabil_KeyDown

    lErro = Grid_Trata_Tecla1(KeyCode, objGridInt)
    If lErro <> SUCESSO Then Error 44063

    Call Trata_SeqContraPartida(objTela.Controls("CTBGridContabil").Row)

    dColunaSoma = Contabil_GridColuna_Soma(iGrid_Debito_Col)
    objTela.Controls("CTBTotalDebito") = Format(dColunaSoma, "Standard")
    dColunaSoma = Contabil_GridColuna_Soma(iGrid_Credito_Col)
    objTela.Controls("CTBTotalCredito") = Format(dColunaSoma, "Standard")

    Exit Sub
    
Erro_Contabil_GridContabil_KeyDown:

    Select Case Err
    
        Case 44063
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147842)
    
    End Select
    
    Exit Sub

End Sub

Sub Contabil_GridContabil_LeaveCell()

    Call objTela.Saida_Celula(objGridInt)

End Sub

Sub Contabil_GridContabil_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridInt)

End Sub

Sub Contabil_GridContabil_RowColChange()

    Call Grid_RowColChange(objGridInt)

End Sub

Sub Contabil_GridContabil_Scroll()

    Call Grid_Scroll(objGridInt)

End Sub

Private Function Verifica_Gera_Lote_Automatico(iGeraLoteAutomatico As Integer) As Long

Dim lErro As Long
Dim objCTBGlobal As New ClassCTBGlobal

On Error GoTo Erro_Verifica_Gera_Lote_Automatico

    If sSiglaModulo = MODULO_CONTASAPAGAR Then
    
        iGeraLoteAutomatico = objCTBGlobal.gobjCTB.iCPGeraLoteAutomatico
        
    ElseIf sSiglaModulo = MODULO_CONTASARECEBER Then
    
        iGeraLoteAutomatico = objCTBGlobal.gobjCTB.iCRGeraLoteAutomatico
        
    ElseIf sSiglaModulo = MODULO_TESOURARIA Then

        iGeraLoteAutomatico = objCTBGlobal.gobjCTB.iTESGeraLoteAutomatico

    ElseIf sSiglaModulo = MODULO_FATURAMENTO Then

        iGeraLoteAutomatico = objCTBGlobal.gobjCTB.iFATGeraLoteAutomatico

    ElseIf sSiglaModulo = MODULO_ESTOQUE Then

        iGeraLoteAutomatico = objCTBGlobal.gobjCTB.iESTGeraLoteAutomatico

    End If

    Verifica_Gera_Lote_Automatico = SUCESSO

    Exit Function

Erro_Verifica_Gera_Lote_Automatico:

    Verifica_Gera_Lote_Automatico = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147843)

    End Select

    Exit Function

End Function

Function Contabil_Gravar_Registro(lNumIntDocOrigem As Long, lCliForn As Long, iFilialCliForn As Integer, dtDataEstoque As Date, Optional iCustoRealProducao As Integer = -1, Optional iErroDocContabilizado As Integer = 1, Optional lNumNotaFiscal As Long = -1, Optional dValorLivroAuxiliar As Double = 0, Optional sDocOrigem As String = "") As Long
'grava os dados contábeis, originados pelo documento lNumIntDocOrigem
'Quando iCustoRealProducao = LANPENDENTE_APROPR_CRPROD indica que os lançamentos serão apropriados pelo custo real de produção
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
'Se a contabilizacao já tiver sido atualizada não será alterada
'iErroDocContabilizado indica se deve ser dada msg de erro se o docto já tiver sua contabilidade lancada (atualizada)
'lNumNotaFiscal se passado como parametro (valor <> -1) indica que todas as insercoes de {NF} na string de historico devem ser substituidas pelo valor passado como parametro
'iSubTipoTransacao indica para qual subtipo da transação está gravando a contabilidade. Antigamente cada tela era associada a uma transação. Agora telas como notas fiscais poderão ser tratados como transações diferentes.
'reprocessamento (incluido dtDataEstoque como parametro)

Dim lErro As Long
Dim colLancamento_Detalhe As New Collection
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim colLancamentos As New Collection
Dim objLanc As ClassLancamento_Detalhe
Dim colLancamento_Detalhe_Custo As New Collection
Dim objLancamento_Cabecalho_Custo As New ClassLancamento_Cabecalho
Dim sModelo As String
Dim iIndice As Integer
Dim iAchou As Integer
Dim iGeraLoteAutomatico As Integer
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim dtData As Date
Dim iFilialEmpresaSalva As Integer
    
On Error GoTo Erro_Contabil_Gravar_Registro

    iFilialEmpresaSalva = giFilialEmpresa

    Do While giFilialEmpresa > 0 And giFilialEmpresa < 100

        If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
            lNumIntDocOrigemExt = lNumIntDocOrigem
            lNumNotaFiscalExt = lNumNotaFiscal
    
            lErro = Contabil_Inicializa_Contabilidade2()
            If lErro <> SUCESSO Then gError 60783
    
            If giContabGerencial = 1 Then
    
                lErro = Contabil_Exclui1(lNumIntDocOrigem, CONTABIL_EXCLUI_EFETIVO, dtData)
                If lErro <> SUCESSO And lErro <> 59220 Then gError 36394
    
            Else
    
                'exclui os dados contábeis pendentes (se houverem) relativos a lNumIntDocOrigem e se tiver dados contabilizados, estorna-os
                lErro = Contabil_Exclui1(lNumIntDocOrigem, CONTABIL_EXCLUI_ESTORNO_DOC_ZERADO, dtData)
                If lErro <> SUCESSO And lErro <> 59220 Then gError 36394
            
            End If
            
            'se o lançamento já estiver contabilizado
            If lErro = 59220 Then
            
                'faz o tratamento dos lançamentos que já tiverem sido contabilizados
                lErro = Trata_Lancamento_Contabilizado(lNumIntDocOrigem, dtData, dtDataEstoque, iErroDocContabilizado)
                If lErro <> SUCESSO And lErro <> 89712 Then gError 89711
                
                If lErro = 89712 Then gError 59254
            
            Else
            
                'se a contabilizacao ainda nao havia sido atualizada
                
                'se a contabilidade não estiver sido alterada pelo usuário ==> carrega a contabilidade padrão (se houver)
                If iContabilidade_Alterada <> CONTABILIDADE_ALTERADA Then
        
                    lErro = Contabil_Carga_Modelo_Padrao_Int()
                    If lErro <> SUCESSO Then gError 36387
                    
                'se o flag de geração de lançamentos automáticos estiver ligado ==> regera os lançamentos automáticos
                ElseIf objTela.Controls("CTBLancAutomatico").Value = MARCADO Then
        
                    Call Contabil_Limpa_GridContabil1
        
                    Call Contabil_Modelo_Click
                    
                End If
        
                'Verifica a existencia de pelo menos um lançamento contabil
                If objGridInt.iLinhasExistentes > 0 Then
            
                    lErro = Contabil_Move_Tela_Memoria(colLancamento_Detalhe, objLancamento_Cabecalho, lCliForn, iFilialCliForn, lNumIntDocOrigem, lNumNotaFiscal, sDocOrigem)
                    If lErro <> SUCESSO Then gError 36382
                    
                    If dValorLivroAuxiliar <> 0 Then
                        
                        If colLancamento_Detalhe.Count > 0 Then
                            Set objLancamento_Detalhe = colLancamento_Detalhe(1)
                            objLancamento_Detalhe.dValorLivroAuxiliar = dValorLivroAuxiliar
                        End If
                                            
                        
                    End If
                    
                    'se for um lançamento dentro de um lote
                    If objLancamento_Cabecalho.iLote > 0 Then
                    
                        'gera os lançamentos pendentes pois estão dentro de um lote
                        lErro = CF("Lancamento_Grava0", objLancamento_Cabecalho, colLancamento_Detalhe)
                        If lErro <> SUCESSO Then gError 36384
                                
                        'verifica se deve gerar numeração de lote automaticamente
                        lErro = Verifica_Gera_Lote_Automatico(iGeraLoteAutomatico)
                        If lErro <> SUCESSO Then gError 36343
                
                        'se é para gerar o lote automatico
                        If iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO Then
                
                            'verifica se o usuario já tem um numero de lote associado
                            lErro = CF("UsuarioLote_Grava_Lote", gsUsuario, sSiglaModulo, objLancamento_Cabecalho.iLote)
                            If lErro <> SUCESSO Then gError 60894
                            
                        End If
                    
                    Else
                    
                        'contabiliza os lançamentos imediatamente
                        lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancamento_Detalhe)
                        If lErro <> SUCESSO Then gError 36386
                    
                    End If
                    
                End If
        
                'se não houver modelo selecionado não há como gerar os lançamentos de custo
                If objTela.Controls("CTBModelo").ListIndex >= 0 Then
        
                    sModelo = objTela.Controls("CTBModelo").Text
        
                    'Verifica se a data contabil está preenchida
                    If Len(objTela.Controls("CTBDataContabil").ClipText) = 0 Then gError 44624
        
                    objLancamento_Cabecalho_Custo.dtData = CDate(objTela.Controls("CTBDataContabil").Text)
        
                    'gera os lançamentos de custo
                    'reprocessamento (incluido dtDataEstoque como parametro)
                    lErro = Contabil_Gera_Lancamento_Custo(colLancamento_Detalhe_Custo, objLancamento_Cabecalho_Custo, lCliForn, iFilialCliForn, lNumIntDocOrigem, sModelo, lNumNotaFiscal, sDocOrigem, iSubTipoTransacaoExt, dtDataEstoque)
                    If lErro <> SUCESSO And lErro <> 96101 Then gError 36411
            
                    If lErro = SUCESSO Then
                    
                        If iCustoRealProducao <> -1 Then
                        
                            If iCustoRealProducao = LANPENDENTE_APROPR_CRPROD Then objLancamento_Cabecalho_Custo.iCustoRealProducao = iCustoRealProducao
                                            
                        End If
                
                        'grava os lançamentos de custo pendentes já que sempre estarão dentro de um lote
                        lErro = CF("Lancamento_Grava0", objLancamento_Cabecalho_Custo, colLancamento_Detalhe_Custo)
                        If lErro <> SUCESSO Then gError 36455
                                
                    End If
    
                End If
    
                'se a contabilidade nao foi informada e o sistema exige contabilizacao para esta transacao ==> erro
                If colLancamento_Detalhe.Count = 0 And colLancamento_Detalhe_Custo.Count = 0 And iObrigatoriedade = EXIGE_CONTABILIZACAO Then gError 110000
    
            End If
        
        End If

        If giContabGerencial = 0 Then Exit Do

        Set colLancamento_Detalhe = New Collection
        Set colLancamento_Detalhe_Custo = New Collection

        giFilialEmpresa = giFilialEmpresa - giFilialAuxiliar
        
    Loop
    
    giFilialEmpresa = iFilialEmpresaSalva

    Contabil_Gravar_Registro = SUCESSO

    Exit Function

Erro_Contabil_Gravar_Registro:

    Contabil_Gravar_Registro = gErr

    giFilialEmpresa = iFilialEmpresaSalva

    Select Case gErr
    
        Case 36387, 36382, 36384, 36386, 36394, 36411, 36455, 60783, 60894, 89711
        
        Case 44624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_CONTABIL_NAO_PREENCHIDA", gErr)

        Case 59254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_CONTABILIZADOS", gErr)
        
        Case 110000
             lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTABILIDADE_OBRIGATORIA", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147844)

    End Select
    
    Exit Function

End Function

Private Function Trata_Lancamento_Contabilizado(lNumIntDocOrigem As Long, dtData As Date, dtDataEstoque As Date, iErroDocContabilizado As Integer) As Long
'faz o tratamento dos lançamentos que já tiverem sido contabilizados

Dim dtData1 As Date
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long
Dim colLancamento_Detalhe As New Collection
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho

On Error GoTo Erro_Trata_Lancamento_Contabilizado

    'se a data estiver preenchida
    If Len(objTela.Controls("CTBDataContabil").ClipText) <> 0 Then

        'se a data de contabilização foi alterada ==> dar aviso perguntando se confirma a alteracao
        If CDate(objTela.Controls("CTBDataContabil").Text) <> dtData Then
        
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_DATA_CONTABIL", CDate(objTela.Controls("CTBDataContabil").Text), dtData)

            If vbMsgRes = vbYes Then
            
                dtData1 = CDate(objTela.Controls("CTBDataContabil").Text)
            
                'Exclui os lançamentos da transação em questão, acerta os totais dos lotes envolvidos e inclui um novo conjunto de lançamentos na nova data
                lErro = CF("Lancamento_Altera_Data", giFilialEmpresa, iCodTransacao, lNumIntDocOrigem, dtData1, dtDataEstoque)
                If lErro <> SUCESSO Then gError 89710
            
            
            ElseIf iErroDocContabilizado Then
            
                gError 89703
                
            End If
    
        ElseIf iErroDocContabilizado Then
            
            gError 89712
        End If
        
    ElseIf iErroDocContabilizado Then
        
        gError 89712
        
    End If
    
    'move os lançamentos da tela para colLancamento_Detalhe e objLancamento_Cabecalho
    lErro = Contabil_Move_Tela_Memoria(colLancamento_Detalhe, objLancamento_Cabecalho, 0, 0, 0, 0, "")
    If lErro <> SUCESSO Then gError 92102

    lErro = CF("Lancamento_Altera_Historico_Trans", objLancamento_Cabecalho, colLancamento_Detalhe)
    If lErro <> SUCESSO Then gError 92103
    
    Trata_Lancamento_Contabilizado = SUCESSO

    Exit Function

Erro_Trata_Lancamento_Contabilizado:

    Trata_Lancamento_Contabilizado = gErr

    Select Case gErr
    
        Case 89703, 89710, 89712, 92102, 92103, 92104, 92105, 92106, 92107
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147845)

    End Select
    
    Exit Function

End Function

Function Contabil_Exclui(lNumIntDocOrigem As Long, Optional iSemEstornar As Integer = 0, Optional dtData As Date, Optional ByVal sOrigem As String, Optional ByVal dtDataEstorno As Date = DATA_NULA, Optional ByVal iFilialEmpresa As Integer = 0) As Long
'exclui os dados contábeis, se houverem, relativos a lNumIntDoc passado como parametro
'o parametro iSemEstornar indica se deve (ou nao) ser estornada contabilizacao já lancada
'se iSemEstornar = 2 significa que só deve estornar se o documento da tela for nulo
'se iSemEstornar = 3 indica que os lançamento já contabilizados devem ser excluidos
'devolve a data no caso do lançamento já tiver sido contabilizado. Esta data servirá para verificar se é um caso de alteração de data.
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lDocParam As Long
Dim iFilialEmpresaSalva As Integer
    
On Error GoTo Erro_Contabil_Exclui

    iFilialEmpresaSalva = giFilialEmpresa

    Do While giFilialEmpresa > 0 And giFilialEmpresa < 100

        lErro = Contabil_Exclui1(lNumIntDocOrigem, iSemEstornar, dtData, sOrigem, dtDataEstorno, iFilialEmpresa, lDocParam)
        If lErro <> SUCESSO Then gError 188056

        If giContabGerencial = 0 Then Exit Do

        giFilialEmpresa = giFilialEmpresa - giFilialAuxiliar
        If iFilialEmpresa <> 0 Then iFilialEmpresa = iFilialEmpresa - giFilialAuxiliar
        
    Loop
    
    giFilialEmpresa = iFilialEmpresaSalva

    Contabil_Exclui = SUCESSO

    Exit Function

Erro_Contabil_Exclui:

    Contabil_Exclui = gErr

    giFilialEmpresa = iFilialEmpresaSalva

    Select Case gErr
    
        Case 188056
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188055)

    End Select
    
    Exit Function

End Function

Private Function Contabil_Move_Tela_Memoria(colLancamento_Detalhe As Collection, objLancamento_Cabecalho As ClassLancamento_Cabecalho, lCliForn As Long, iFilialCliForn As Integer, lNumIntDocOrigem As Long, lNumNotaFiscal As Long, sDocOrigem As String) As Long
    
Dim lErro As Long
Dim lDoc As Long
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim iIndice1 As Integer
Dim dSoma As Double
Dim iPeriodoDoc As Integer
Dim iExercicioDoc As Integer
Dim objPeriodo As New ClassPeriodo
    
On Error GoTo Erro_Contabil_Move_Tela_Memoria
    
    'Verifica se a data contabil está preenchida
    If Len(objTela.Controls("CTBDataContabil").ClipText) = 0 Then Error 36372

    'Preenche Objeto Lançamento_Cabeçalho
    objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
    objLancamento_Cabecalho.sOrigem = sSiglaModulo
    objLancamento_Cabecalho.dtData = CDate(objTela.Controls("CTBDataContabil").Text)
    objLancamento_Cabecalho.lNumIntDoc = lNumIntDocOrigem
    objLancamento_Cabecalho.iTransacao = iCodTransacao
    objLancamento_Cabecalho.sModelo = objTela.Controls("CTBModelo").Text
    objLancamento_Cabecalho.iSubTipo = iSubTipoTransacaoExt
    
    lErro = CF("Periodo_Le1", objLancamento_Cabecalho.dtData, objPeriodo, objLancamento_Cabecalho.iFilialEmpresa)
    If lErro <> SUCESSO Then Error 20819
    
    objLancamento_Cabecalho.iExercicio = objPeriodo.iExercicio
    objLancamento_Cabecalho.iPeriodoLan = objPeriodo.iPeriodo
    
    If objLancamento_Cabecalho.lDoc = 0 Then
        'Se o Documento contábil não estiver preenchido ==> gera um doc automatico
        If Len(objTela.Controls("CTBDocumento").ClipText) = 0 Then
        
        
            lErro = CF("Voucher_Automatico_Trans", objLancamento_Cabecalho.iFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, objLancamento_Cabecalho.sOrigem, lDoc)
            If lErro <> SUCESSO Then Error 20820
            
            objLancamento_Cabecalho.lDoc = lDoc
        
        Else
        
            objLancamento_Cabecalho.lDoc = CLng(objTela.Controls("CTBDocumento").ClipText)
        
        End If
    
    End If
    
    'se o lote estiver preenchido ==> não pode ser com o valor zero
'    If Len(objTela.Controls("CTBLote").ClipText) > 0 And giTipoVersao = VERSAO_FULL Then
    If Len(objTela.Controls("CTBLote").ClipText) > 0 Then
        objLancamento_Cabecalho.iLote = CInt(objTela.Controls("CTBLote").ClipText)
    Else
        'se não estiver preenchido o lote ==> atualizacao imediata e o valor do lote será zero internamente
        objLancamento_Cabecalho.iLote = 0
    End If

    'Preenche Objeto Lançamento_Detalhe
    lErro = Contabil_Grid_Lancamento_Detalhe(colLancamento_Detalhe, lCliForn, iFilialCliForn, lNumIntDocOrigem, lNumNotaFiscal, sDocOrigem)
    If lErro <> SUCESSO Then Error 36380
    
    'Testa se soma dos débitos é igual a soma dos créditos
    dSoma = 0

    For Each objLancamento_Detalhe In colLancamento_Detalhe
        dSoma = dSoma + objLancamento_Detalhe.dValor
    Next

    dSoma = Format(dSoma, "Fixed")

    If Abs(dSoma) > DELTA_VALORMONETARIO Then Error 36381
            
    Contabil_Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Contabil_Move_Tela_Memoria:

    Contabil_Move_Tela_Memoria = Err

    Select Case Err
    
        Case 20819, 20820, 36380, 36382
        
        Case 36372
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_CONTABIL_NAO_PREENCHIDA", Err)
    
        Case 36373
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_CONTABIL_NAO_PREENCHIDO", Err)
    
        Case 36374
        
        Case 36381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_BALANCEADO", Err, objLancamento_Cabecalho.lDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147847)
            
    End Select
    
    Exit Function
    
End Function

Private Function Contabil_Gera_Lancamento_Custo(colLancamento_Detalhe As Collection, objLancamento_Cabecalho As ClassLancamento_Cabecalho, lCliForn As Long, iFilialCliForn As Integer, lNumIntDocOrigem As Long, sModelo As String, lNumNotaFiscal As Long, sDocOrigem As String, iSubTipoTransacao As Integer, dtDataEstoque As Date) As Long
'gera os lançamentos contábeis relacionados a custo.
'reprocessamento (incluido dtDataEstoque como parametro)

Dim lErro As Long
Dim lDoc As Long
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim iIndice1 As Integer
Dim dSoma As Double
Dim iPeriodoDoc As Integer
Dim iExercicioDoc As Integer
Dim objLote As New ClassLote
Dim objPeriodo As New ClassPeriodo


On Error GoTo Erro_Contabil_Gera_Lancamento_Custo
    
    'Preenche Objeto Lançamento_Cabeçalho
    objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
    objLancamento_Cabecalho.sOrigem = sOrigemCusto
    
    'Coloca o periodo relativo a data na tela
    lErro = CF("Periodo_Le", objLancamento_Cabecalho.dtData, objPeriodo)
    If lErro <> SUCESSO Then gError 96099
    
    'Preenche Objeto Lançamento_Detalhe
    'reprocessamento (incluido dtDataEstoque como parametro)
    lErro = Contabil_Grid_Custo_Detalhe(colLancamento_Detalhe, lCliForn, iFilialCliForn, lNumIntDocOrigem, objLote, objLancamento_Cabecalho, sModelo, lNumNotaFiscal, sDocOrigem, iSubTipoTransacao, dtDataEstoque)
    If lErro <> SUCESSO And lErro <> 36775 Then gError 96100
        
    If lErro <> SUCESSO Then gError 96101
    
    objLote.iFilialEmpresa = giFilialEmpresa
    objLote.sOrigem = sOrigemCusto
    objLote.iExercicio = objPeriodo.iExercicio
    objLote.iPeriodo = objPeriodo.iPeriodo
    
    If objLancamento_Cabecalho.iLote = 0 Or giContabGerencial = 0 Then
    
        lErro = Contabil_Retorna_Lote_Custo(objLote)
        If lErro <> SUCESSO Then gError 96102
        
        objLancamento_Cabecalho.iLote = objLote.iLote
    
    End If
    
    If objLancamento_Cabecalho.lDoc = 0 Or giContabGerencial = 0 Then
    
        lErro = CF("Voucher_Automatico1", giFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, sOrigemCusto, lDoc)
        If lErro <> SUCESSO Then gError 96103
                
        objLancamento_Cabecalho.lDoc = lDoc
    
    End If
    
    Contabil_Gera_Lancamento_Custo = SUCESSO

    Exit Function

Erro_Contabil_Gera_Lancamento_Custo:

    Contabil_Gera_Lancamento_Custo = gErr

    Select Case gErr
    
        Case 96100, 96102, 96103, 96101, 96099
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147848)
            
    End Select
    
    Exit Function
    
End Function

Private Function Contabil_Grid_Lancamento_Detalhe(colLancamento_Detalhe As Collection, lCliForn As Long, iFilialCliForn As Integer, lNumIntDocOrigem As Long, lNumNotaFiscal As Long, sDocOrigem As String, Optional ByVal bConsulta As Boolean = False) As Long

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
Dim iLinha As Integer
Dim sHistorico As String
Dim iPosicao As Integer

On Error GoTo Erro_Contabil_Grid_Lancamento_Detalhe

    For iIndice1 = 1 To objGridInt.iLinhasExistentes
        
        Set objLancamento_Detalhe = New ClassLancamento_Detalhe
        
        objLancamento_Detalhe.lCliForn = lCliForn
        objLancamento_Detalhe.iFilialCliForn = iFilialCliForn
        objLancamento_Detalhe.lNumIntDoc = lNumIntDocOrigem
        objLancamento_Detalhe.iTransacao = iCodTransacao
        objLancamento_Detalhe.iClienteFornecedor = iClienteFornecedor
        objLancamento_Detalhe.sDocOrigem = sDocOrigem
        If Not bConsulta Then objLancamento_Detalhe.sModelo = objTela.Controls("CTBModelo").Text
                
        'Testa para ver se houve crédito ou débito
        If Len(objTela.Controls("CTBGridContabil").TextMatrix(iIndice1, iGrid_Credito_Col)) > 0 Then
            dValorCredito = CDbl(objTela.Controls("CTBGridContabil").TextMatrix(iIndice1, iGrid_Credito_Col))
        Else
            dValorCredito = 0
        End If
            
        If Len(objTela.Controls("CTBGridContabil").TextMatrix(iIndice1, iGrid_Debito_Col)) > 0 Then
            dValorDebito = CDbl(objTela.Controls("CTBGridContabil").TextMatrix(iIndice1, iGrid_Debito_Col))
        Else
            dValorDebito = 0
        End If
    
        'só guarda os lançamentos que tem valor
        If dValorDebito <> 0 Or dValorCredito <> 0 Then
            
            iLinha = iLinha + 1
            
            sConta = objTela.Controls("CTBGridContabil").TextMatrix(iIndice1, iGrid_Conta_Col)
                
            If Len(Trim(sConta)) = 0 Then Error 36375
                
            objTela.Controls("CTBConta").PromptInclude = False
            objTela.Controls("CTBConta").Text = sConta
            objTela.Controls("CTBConta").PromptInclude = True
                
            lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then Error 36376
            
            If iContaPreenchida = CONTA_VAZIA Then Error 36377
                
            objLancamento_Detalhe.sConta = sContaFormatada
    
            objLancamento_Detalhe.iSeq = iLinha 'a linha que realmente vai ficar armazenada
            objLancamento_Detalhe.iSeqGrid = iIndice1 'a linha original do grid
            
            objLancamento_Detalhe.dValor = dValorCredito - dValorDebito
        
            'armazena o sequencial de contra partida, se estiver preenchido
            If Len(objTela.Controls("CTBGridContabil").TextMatrix(iIndice1, iGrid_SeqContraPartida_Col)) > 0 Then
                objLancamento_Detalhe.iSeqContraPartida = CInt(objTela.Controls("CTBGridContabil").TextMatrix(iIndice1, iGrid_SeqContraPartida_Col))
                
                lErro = Armazena_Contra_Partida(colContraPartida, objLancamento_Detalhe)
                If lErro <> SUCESSO Then Error 44057
                
            End If
        
            'Armazena Histórico e Ccl
            objLancamento_Detalhe.sHistorico = objTela.Controls("CTBGridContabil").TextMatrix(iIndice1, iGrid_Historico_Col)
            
            
            'se foi passado como parametro o numero da nota fiscal ==> pesquisa por {NF} nos historicos e substitui pelo numero da nota, se houver
            If lNumNotaFiscal <> -1 Then
            
                sHistorico = objLancamento_Detalhe.sHistorico
            
                Call Historico_Troca_Num_Doc(sHistorico, lNumNotaFiscal)
                
                objLancamento_Detalhe.sHistorico = sHistorico
                
            End If
            
            objLancamento_Detalhe.sProduto = ""
                
            'verifica se o historico tem parametros que deveriam ter sido substituidos
            If InStr(objLancamento_Detalhe.sHistorico, CARACTER_HISTORICO_PARAM) <> 0 Then Error 44082
                
            'Se está usando Centro de Custo/Lucro, armazena-o
            If iGrid_Ccl_Col <> 999 Then
                    
                sCcl = objTela.Controls("CTBGridContabil").TextMatrix(iIndice1, iGrid_Ccl_Col)
                
                lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
                If lErro <> SUCESSO Then Error 36379
                
                If iCclPreenchida = CCL_PREENCHIDA Then
                    objLancamento_Detalhe.sCcl = sCclFormatada
                Else
                    objLancamento_Detalhe.sCcl = ""
                End If
                    
            End If
                    
            'Armazena a indicação de aglutinação
            objLancamento_Detalhe.iAglutina = StrParaInt(objTela.Controls("CTBGridContabil").TextMatrix(iIndice1, iGrid_Aglutina_Col))
                    
            If giContabGerencial = 1 Then
                objLancamento_Detalhe.iGerencial = StrParaInt(objTela.Controls("CTBGridContabil").TextMatrix(iIndice1, iGrid_Gerencial_Col))
                If (objLancamento_Detalhe.iGerencial = 0 Or objLancamento_Detalhe.iGerencial = 2) And ((giFilialEmpresa > giFilialAuxiliar And giFilialAuxiliar > 0) Or (giFilialEmpresa < Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                    'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                    colLancamento_Detalhe.Add objLancamento_Detalhe
                End If
                    
                If (objLancamento_Detalhe.iGerencial = 1 Or objLancamento_Detalhe.iGerencial = 2) And ((giFilialEmpresa < giFilialAuxiliar And giFilialAuxiliar > 0) Or (giFilialEmpresa > Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                    colLancamento_Detalhe.Add objLancamento_Detalhe
                End If
                
            Else
                    
                'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                colLancamento_Detalhe.Add objLancamento_Detalhe
            End If
            
        
        End If
        
    Next
    
    'acerta o sequencial de contra partida pois alguns elementos do grid podem não ser armazenados por estarem zerados
    lErro = Acerta_Contra_Partida(colLancamento_Detalhe)
    If lErro <> SUCESSO Then Error 44625
    
    lErro = Testa_Contra_Partida(colLancamento_Detalhe, colContraPartida)
    If lErro <> SUCESSO Then Error 44058
    
    Contabil_Grid_Lancamento_Detalhe = SUCESSO

    Exit Function

Erro_Contabil_Grid_Lancamento_Detalhe:

    Contabil_Grid_Lancamento_Detalhe = Err

    Select Case Err
    
        Case 36375, 36377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)
                
        Case 36376, 36379, 44057, 44058, 44625
        
        Case 44082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTORICO_PARAM", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147849)
            
    End Select
    
    Exit Function

End Function

Private Function Contabil_Grid_Custo_Detalhe(colLancamento_Detalhe As Collection, lCliForn As Long, iFilialCliForn As Integer, lNumIntDocOrigem As Long, objLote As ClassLote, objLancamento_Cabecalho As ClassLancamento_Cabecalho, sModelo As String, lNumNotaFiscal As Long, sDocOrigem As String, iSubTipoTransacao As Integer, dtDataEstoque As Date) As Long
'reprocessamento (incluido dtDataEstoque como parametro)

Dim iIndice1 As Integer
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim sConta As String
Dim sContaFormatada As String
Dim sCcl As String
Dim sCclFormatada As String
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim lErro As Long
Dim colLancamentos As New ClassColLancamentos
Dim objLancamento As ClassLancamentos
Dim colProdutos As New Collection
Dim colTotais As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim colContraPartida As New Collection
Dim iLinha As Integer
Dim sHistorico As String
Dim iPosicao As Integer

On Error GoTo Erro_Contabil_Grid_Custo_Detalhe

    lErro = CF("Gera_Lancamentos_Custo", sSiglaModulo, objTela, sModelo, colLancamentos, iSubTipoTransacao)
    If lErro <> SUCESSO And lErro <> 36774 Then Error 36446

    If lErro <> SUCESSO Then Error 36775

    iIndice1 = 0
    
    For Each objLancamento In colLancamentos
    
        iIndice1 = iIndice1 + 1
    
        Set objLancamento_Detalhe = New ClassLancamento_Detalhe
        
        objLancamento_Detalhe.lCliForn = lCliForn
        objLancamento_Detalhe.iFilialCliForn = iFilialCliForn
        objLancamento_Detalhe.lNumIntDoc = lNumIntDocOrigem
        objLancamento_Detalhe.iTransacao = iCodTransacao
        objLancamento_Detalhe.iClienteFornecedor = iClienteFornecedor
        objLancamento_Detalhe.sDocOrigem = sDocOrigem
        objLancamento_Detalhe.sModelo = sModelo
                
        'se o lançamento tiver alguma quantidade em crédito ou débito
        If objLancamento.dDebito <> 0 Or objLancamento.dCredito <> 0 Then
        
            iLinha = iLinha + 1
            
            sConta = objLancamento.sConta
                
            If Len(Trim(sConta)) = 0 Then Error 36448
                
            objTela.Controls("CTBConta").PromptInclude = False
            objTela.Controls("CTBConta").Text = sConta
            objTela.Controls("CTBConta").PromptInclude = True
                
            lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then Error 36449
                
            objLancamento_Detalhe.sConta = sContaFormatada
          
            objLancamento_Detalhe.iSeq = iLinha 'a linha que realmente vai ficar armazenada
            objLancamento_Detalhe.iSeqGrid = iIndice1 'a linha original do grid
        
            If objLancamento.dDebito <> 0 And objLancamento.dCredito <> 0 Then Error 36451
                
            objLancamento_Detalhe.dValor = objLancamento.dCredito - objLancamento.dDebito
            
            objLote.dTotCre = objLote.dTotCre + objLancamento.dCredito
            objLote.dTotDeb = objLote.dTotDeb + objLancamento.dDebito
            objLote.iNumLancAtual = objLote.iNumLancAtual + 1
            
            'armazena o sequencial de contra partida, se estiver preenchido
            If objLancamento.iSeqContraPartida <> 0 Then
            
                objLancamento_Detalhe.iSeqContraPartida = objLancamento.iSeqContraPartida
                
                lErro = Armazena_Contra_Partida(colContraPartida, objLancamento_Detalhe)
                If lErro <> SUCESSO Then Error 44059
                
            End If
            
            'Armazena Histórico e Ccl
            objLancamento_Detalhe.sHistorico = objLancamento.sHistorico
                
            'verifica se o historico tem parametros que deveriam ter sido substituidos
            If InStr(objLancamento_Detalhe.sHistorico, CARACTER_HISTORICO_PARAM) <> 0 Then Error 44083
            
            'se foi passado como parametro o numero da nota fiscal ==> pesquisa por {NF} nos historicos e substitui pelo numero da nota, se houver
            If lNumNotaFiscal <> -1 Then
            
                sHistorico = objLancamento_Detalhe.sHistorico
            
                Call Historico_Troca_Num_Doc(sHistorico, lNumNotaFiscal)
                                
                objLancamento_Detalhe.sHistorico = sHistorico
                
            End If
            
            
            'Se está usando Centro de Custo/Lucro, armazena-o
            If iGrid_Ccl_Col <> 999 Then
                    
                sCcl = objLancamento.sCcl
                
                lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
                If lErro <> SUCESSO Then Error 36452
                
                If iCclPreenchida = CCL_PREENCHIDA Then
                    objLancamento_Detalhe.sCcl = sCclFormatada
                Else
                    objLancamento_Detalhe.sCcl = ""
                End If
                    
            End If
                    
            'Armazena a indicação de aglutinação
            objLancamento_Detalhe.iAglutina = objLancamento.iAglutina
            
            lErro = CF("Produto_Formata", objLancamento.sProduto, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then Error 36811
            
            'Armazena a indicação de produto
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
                objLancamento_Detalhe.sProduto = sProdutoFormatado
                'reprocessamento
                objLancamento_Detalhe.dQuantidade = objLancamento_Detalhe.dValor
                objLancamento_Detalhe.dtDataEstoque = dtDataEstoque
            
                'armazena os totais das quantidade de cada produto.
                lErro = Contabil_Armazena_Produtos_Totais(objLancamento_Detalhe, objLancamento, colProdutos, colTotais)
                If lErro <> SUCESSO Then Error 36453
                
            Else
            
               objLancamento_Detalhe.sProduto = ""
               
            End If
    
            'Armazena a indicação de qual escaninho o custo vai ser lido
            objLancamento_Detalhe.iEscaninho_Custo = objLancamento.iEscaninho_Custo
            
            If giContabGerencial = 1 Then
                objLancamento_Detalhe.iGerencial = objLancamento.iGerencial
                If (objLancamento_Detalhe.iGerencial = 0 Or objLancamento_Detalhe.iGerencial = 2) And ((giFilialEmpresa > giFilialAuxiliar And giFilialAuxiliar > 0) Or (giFilialEmpresa < Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                    'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                    colLancamento_Detalhe.Add objLancamento_Detalhe
                End If
                    
                If (objLancamento_Detalhe.iGerencial = 1 Or objLancamento_Detalhe.iGerencial = 2) And ((giFilialEmpresa < giFilialAuxiliar And giFilialAuxiliar > 0) Or (giFilialEmpresa > Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                    colLancamento_Detalhe.Add objLancamento_Detalhe
                End If
                
            Else
                    
                'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                colLancamento_Detalhe.Add objLancamento_Detalhe
            End If
        
        End If
    Next
    
    'testa se os totais de todos os produtos dá zero. Se não der há um desbalanceamento = erro.
    lErro = Contabil_Testa_Produtos_Totais(colProdutos, colTotais)
    If lErro <> SUCESSO Then Error 36454
    
    'acerta o sequencial de contra partida pois alguns elementos do grid podem não ser armazenados por estarem zerados
    lErro = Acerta_Contra_Partida(colLancamento_Detalhe)
    If lErro <> SUCESSO Then Error 44626
    
    lErro = Testa_Contra_Partida(colLancamento_Detalhe, colContraPartida)
    If lErro <> SUCESSO Then Error 44060
    
    Contabil_Grid_Custo_Detalhe = SUCESSO

    Exit Function

Erro_Contabil_Grid_Custo_Detalhe:

    Contabil_Grid_Custo_Detalhe = Err

    Select Case Err
    
        Case 36446
    
        Case 36448
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANC_CUSTO_CONTA_NAO_INFORMADA", Err)
                
        Case 36449, 36452, 36453, 36454, 36775, 36811, 44059, 44060, 44626
        
        Case 36450
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANC_CUSTO_QUANT_NAO_PREENCHIDA", Err)
            
        Case 36451
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANC_CUSTO_DEBCRE_PREENCHIDOS", Err)
            
        Case 44083
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTORICO_PARAM", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147850)
            
    End Select
    
    Exit Function

End Function

Private Function Contabil_Testa_Produtos_Totais(colProdutos As Collection, colTotais As Collection) As Long
'testa se os totais de todos os produtos dá zero. Se não der há um desbalanceamento = erro.

Dim vProduto As Variant
Dim vTotal As Variant
Dim iProduto As Integer
Dim iAchou As Integer
Dim lErro As Long

On Error GoTo Erro_Contabil_Testa_Produtos_Totais

    For Each vTotal In colTotais
                
        iProduto = iProduto + 1
        If Abs(vTotal) > 0.000001 Then Error 36447
        
    Next
            
    Contabil_Testa_Produtos_Totais = SUCESSO

    Exit Function

Erro_Contabil_Testa_Produtos_Totais:

    Contabil_Testa_Produtos_Totais = Err

    Select Case Err
    
        Case 36447
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_PRODUTO_DESBALANCEADO", Err, colTotais.Item(iProduto))
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147851)
            
    End Select
    
    Exit Function

End Function

Private Function Contabil_Armazena_Produtos_Totais(objLancamento_Detalhe As ClassLancamento_Detalhe, objLancamento As ClassLancamentos, colProdutos As Collection, colTotais As Collection) As Long
'armazena os totais das quantidade de cada produto.

Dim vProduto As Variant
Dim iProduto As Integer
Dim iAchou As Integer
Dim lErro As Long
Dim vTotal As Variant

On Error GoTo Erro_Contabil_Armazena_Produtos_Totais

    For Each vProduto In colProdutos
        
        iProduto = iProduto + 1
        
        If vProduto = objLancamento.sProduto Then
        
                    
            colTotais.Add CVar(colTotais.Item(iProduto) + objLancamento_Detalhe.dValor), , , iProduto
            colTotais.Remove (iProduto)
            
            iAchou = 1
            Exit For
            
        End If
        
    Next
    
    If iAchou = 0 Then
        
        colProdutos.Add objLancamento.sProduto
        colTotais.Add objLancamento_Detalhe.dValor
        
    End If
            
    Contabil_Armazena_Produtos_Totais = SUCESSO

    Exit Function

Erro_Contabil_Armazena_Produtos_Totais:

    Contabil_Armazena_Produtos_Totais = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147852)
            
    End Select
    
    Exit Function

End Function

Private Function Contabil_Retorna_Lote_Custo(objLote As ClassLote) As Long
'retorna o lote de custo para a Filial/Origem/Exercicio/Periodo em questão. O lote pode já existir ou ser criado neste momento.

Dim lErro As Long

On Error GoTo Erro_Contabil_Retorna_Lote_Custo

    lErro = CF("LotePendente_Le2", objLote)
    If lErro <> SUCESSO And lErro <> 36417 Then Error 36413
    
    'se não encontrou um lote de custo pendente para o periodo em questão ==> criar um
    If lErro = 36417 Then
    
        lErro = CF("LotePendente_Insere_LoteCusto", objLote)
        If lErro <> SUCESSO Then Error 36433
    
    End If
    
    Contabil_Retorna_Lote_Custo = SUCESSO
    
    Exit Function
    
Erro_Contabil_Retorna_Lote_Custo:

    Contabil_Retorna_Lote_Custo = Err
    
    Select Case Err
    
        Case 36413, 36433
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147853)
            
    End Select
    
    Exit Function
    
End Function

Sub Contabil_UpDown_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_Contabil_UpDown_DownClick

    objTela.Controls("CTBDataContabil").SetFocus

    If Len(objTela.Controls("CTBDataContabil").ClipText) > 0 Then

        sData = objTela.Controls("CTBDataContabil").Text
        
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 36468
        
        objTela.Controls("CTBDataContabil").Text = sData
        
    End If
    
    Exit Sub
    
Erro_Contabil_UpDown_DownClick:
    
    Select Case Err
    
        Case 36468
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147854)
        
    End Select
    
    Exit Sub

End Sub

Sub Contabil_UpDown_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_Contabil_UpDown_UpClick

    objTela.Controls("CTBDataContabil").SetFocus

    If Len(objTela.Controls("CTBDataContabil").ClipText) > 0 Then

        sData = objTela.Controls("CTBDataContabil").Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 36467
        
        objTela.Controls("CTBDataContabil").Text = sData
        
    End If
    
    Exit Sub
    
Erro_Contabil_UpDown_UpClick:
    
    Select Case Err
    
        Case 36467
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147855)
        
    End Select
    
    Exit Sub

End Sub

Function Contabil_Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim objMnemonico As New ClassMnemonicoCTBValor
Dim dNum As Double

On Error GoTo Erro_Contabil_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
    
        Case "Arredonda_Num"
            
            objMnemonicoValor.colValor.Add Round(objMnemonicoValor.vParam(1), objMnemonicoValor.vParam(2))
    
        Case Else
    
            objMnemonico.sMnemonico = objMnemonicoValor.sMnemonico
        
            lErro = CF("MnemonicoCTBValor_Le", objMnemonico)
            If lErro <> SUCESSO And lErro <> 39690 Then Error 39691
            
            If lErro = 39690 Then Error 39692
            
            objMnemonicoValor.colValor.Add objMnemonico.sValor
            
    End Select
    
    Contabil_Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Contabil_Calcula_Mnemonico:

    Contabil_Calcula_Mnemonico = Err

    Select Case Err
        
        Case 39691
        
        Case 39692
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CALCULO_MNEMONICO_INEXISTENTE", Err, objMnemonicoValor.sMnemonico)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147856)

    End Select

    Exit Function

End Function

Public Function Contabil_Inicializa_Contabilidade1(objTela1 As Object, sSiglaModulo1 As String, sTransacao As String, Optional iSubTipoTransacao As Integer = 0, Optional iTransacao As Integer = 0) As Long
'usado na inicialização de objContabil quando estiver trabalhando com transações que não tenham tabstrip de contabilização

Dim objTransacaoCTB As New ClassTransacaoCTB
Dim lErro As Long

On Error GoTo Erro_Contabil_Inicializa_Contabilidade1

    iGeraModelo = 0

    Set objTela = objTela1
    
    iSubTipoTransacaoExt = iSubTipoTransacao

    sSiglaModulo = sSiglaModulo1

    sNomeModulo = gcolModulo.Nome(sSiglaModulo)
    
    sOrigemCusto = gcolModulo.sOrigemCusto(sSiglaModulo)
    
    If iTransacao = 0 Then
        
        objTransacaoCTB.sSigla = sSiglaModulo
        objTransacaoCTB.sTransacao = sTransacao
        objTransacaoCTB.iSubTipo = iSubTipoTransacaoExt
        
        lErro = CF("TransacaoCTB_Le", objTransacaoCTB)
        If lErro <> SUCESSO Then Error 36595
    
        iCodTransacao = objTransacaoCTB.iCodigo
        iClienteFornecedor = objTransacaoCTB.iCliFornecedor
        
       
        iObrigatoriedade = objTransacaoCTB.iObrigatoriedade
    
    Else
    
        iCodTransacao = iTransacao
        
    End If
    
    iContabilidade_Alterada = 0
    
    Contabil_Inicializa_Contabilidade1 = SUCESSO

    Exit Function

Erro_Contabil_Inicializa_Contabilidade1:

    Contabil_Inicializa_Contabilidade1 = Err

    Select Case Err

        Case 36595

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147857)

    End Select

    Exit Function

End Function

Function Contabil_Gravar_Registro1(lNumIntDocOrigem As Long, lCliForn As Long, iFilialCliForn As Integer, dtDataEstoque As Date, objLancamento_Cabecalho As ClassLancamento_Cabecalho, sModelo As String, Optional lNumNotaFiscal As Long = -1, Optional dValorLivroAuxiliar As Double = 0, Optional sDocOrigem As String = "") As Long
'usada para contabilizacao automatica
'grava os dados contábeis, originados pelo documento lNumIntDocOrigem
'os lançamentos são gerados a partir do modelo passado como parametro (sModelo)
'Em objLancamento_Cabecalho devem estar preenchidos os campos: iFilialEmpresa, dtData, iCustoRealProducao, iLote (se lancamento em lote), doc
'Quando iCustoRealProducao = LANPENDENTE_APROPR_CRPROD indica que os lançamentos serão apropriados pelo custo real de produção
'lNumNotaFiscal se passado como parametro (valor <> -1) indica que todas as insercoes de {NF} na string de historico devem ser substituidas pelo valor passado como parametro
'reprocessamento (incluido dtDataEstoque como parametro)
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
    
Dim lErro As Long, sHistorico As String, iPosicao As Integer
Dim colLancamento_Detalhe As New Collection, objLancamento_Detalhe As ClassLancamento_Detalhe
Dim objLanc As ClassLancamento_Detalhe
Dim colLancamento_Detalhe_Custo As New Collection
Dim objLancamento_Cabecalho_Custo As New ClassLancamento_Cabecalho
Dim dtData As Date
Dim iFilialEmpresaSalva As Integer
Dim iFilialEmpresaSalva1 As Integer
Dim objLote As New ClassLote
Dim objPeriodo As New ClassPeriodo
    
On Error GoTo Erro_Contabil_Gravar_Registro1

    iFilialEmpresaSalva = giFilialEmpresa
    iFilialEmpresaSalva1 = objLancamento_Cabecalho.iFilialEmpresa

    Do While giFilialEmpresa > 0 And giFilialEmpresa < 100
    
        If objLancamento_Cabecalho.lDoc = 0 Then gError 36609
    
        If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
            If giContabGerencial = 1 Then
            
                'exclui os dados contábeis (se houverem) relativos a lNumIntDocOrigem
                lErro = Contabil_Exclui1(lNumIntDocOrigem, CONTABIL_EXCLUI_EFETIVO, dtData)
                If lErro <> SUCESSO And lErro <> 59220 Then gError 36603
    
            Else
    
                'exclui os dados contábeis (se houverem) relativos a lNumIntDocOrigem
                lErro = Contabil_Exclui1(lNumIntDocOrigem, CONTABIL_EXCLUI_SEM_ESTORNO, dtData)
                If lErro <> SUCESSO And lErro <> 59220 Then gError 36603
    
            End If
    
            'se o lançamento já estiver contabilizado
            If lErro = 59220 Then
            
                'faz o tratamento dos lançamentos que já tiverem sido contabilizados
                lErro = Trata_Lancamento_Contabilizado(lNumIntDocOrigem, dtData, dtDataEstoque, AVISA_LANCAMENTOS_CONTABILIZADOS)
                If lErro <> SUCESSO And lErro <> 89712 Then gError 89713
                
                If lErro = 89712 Then gError 59221
            
            Else
            
                'Gera os lançamentos a partir do modelo sModelo e coloca os dados em colLancamento_Detalhe
                lErro = Contabil_Gera_Lancamento_Detalhe(colLancamento_Detalhe, lCliForn, iFilialCliForn, lNumIntDocOrigem, objLancamento_Cabecalho, sModelo, sDocOrigem, iSubTipoTransacaoExt)
                If lErro <> SUCESSO Then gError 36604
                    
                'se foi passado como parametro o numero da nota fiscal ==> pesquisa por {NF} nos historicos e substitui pelo numero da nota, se houver
                If lNumNotaFiscal <> -1 Then
                
                    For Each objLancamento_Detalhe In colLancamento_Detalhe
                    
                        sHistorico = objLancamento_Detalhe.sHistorico
                    
                        Call Historico_Troca_Num_Doc(sHistorico, lNumNotaFiscal)
                                        
                        objLancamento_Detalhe.sHistorico = sHistorico
                    
                    Next
                                
                End If
                
                If dValorLivroAuxiliar <> 0 Then
                
                    If colLancamento_Detalhe.Count > 0 Then
                        Set objLancamento_Detalhe = colLancamento_Detalhe(1)
                        objLancamento_Detalhe.dValorLivroAuxiliar = dValorLivroAuxiliar
                    End If
                    
                End If
                
                'se for um lançamento dentro de um lote
                If objLancamento_Cabecalho.iLote > 0 Then
                
                    If giContabGerencial = 1 Then
                        
                        
                        lErro = CF("Periodo_Le1", objLancamento_Cabecalho.dtData, objPeriodo, objLancamento_Cabecalho.iFilialEmpresa)
                        If lErro <> SUCESSO Then gError 188069
                        
                        objLote.sOrigem = objLancamento_Cabecalho.sOrigem
                        objLote.iExercicio = objPeriodo.iExercicio
                        objLote.iPeriodo = objPeriodo.iPeriodo
                        objLote.iLote = objLancamento_Cabecalho.iLote
                        objLote.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
                    
                        objLote.iStatus = LOTE_DESATUALIZADO
                        lErro = CF("LotePendente_Grava_Trans", objLote)
                        If lErro <> SUCESSO Then gError 188068
    
                    End If
    
                
                    'gera os lançamentos pendentes pois estão dentro de um lote
                    lErro = CF("Lancamento_Grava0", objLancamento_Cabecalho, colLancamento_Detalhe)
                    If lErro <> SUCESSO Then gError 36605
                            
                Else
                
                    'contabiliza os lançamentos imediatamente
                    lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancamento_Detalhe)
                    If lErro <> SUCESSO Then gError 36606
                
                End If
        
                objLancamento_Cabecalho_Custo.dtData = objLancamento_Cabecalho.dtData
            
                'gera os lançamentos de custo
                'reprocessamento (incluido dtDataEstoque como parametro)
                lErro = Contabil_Gera_Lancamento_Custo(colLancamento_Detalhe_Custo, objLancamento_Cabecalho_Custo, lCliForn, iFilialCliForn, lNumIntDocOrigem, sModelo, lNumNotaFiscal, sDocOrigem, iSubTipoTransacaoExt, dtDataEstoque)
                If lErro <> SUCESSO And lErro <> 96101 Then gError 36607
        
                If lErro = SUCESSO Then
                
                    objLancamento_Cabecalho_Custo.iCustoRealProducao = objLancamento_Cabecalho.iCustoRealProducao
        
                    'grava os lançamentos de custo pendentes já que sempre estarão dentro de um lote
                    lErro = CF("Lancamento_Grava0", objLancamento_Cabecalho_Custo, colLancamento_Detalhe_Custo)
                    If lErro <> SUCESSO Then gError 36608
                    
                End If
                
                'se a contabilidade nao foi informada e o sistema exige contabilizacao para esta transacao ==> erro
                If colLancamento_Detalhe.Count = 0 And colLancamento_Detalhe_Custo.Count = 0 And iObrigatoriedade = EXIGE_CONTABILIZACAO Then gError 110001
                
            End If
    
        End If

        If giContabGerencial = 0 Then Exit Do

        giFilialEmpresa = giFilialEmpresa - giFilialAuxiliar
        objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa - giFilialAuxiliar

        Set colLancamento_Detalhe = New Collection
        Set colLancamento_Detalhe_Custo = New Collection
    
    Loop

    giFilialEmpresa = iFilialEmpresaSalva
    objLancamento_Cabecalho.iFilialEmpresa = iFilialEmpresaSalva1

    Contabil_Gravar_Registro1 = SUCESSO

    Exit Function

Erro_Contabil_Gravar_Registro1:

    Contabil_Gravar_Registro1 = gErr

    giFilialEmpresa = iFilialEmpresaSalva
    objLancamento_Cabecalho.iFilialEmpresa = iFilialEmpresaSalva1

    Select Case gErr
    
        Case 36603, 36604, 36605, 36606, 36607, 36608, 89713, 188068, 188069
        
        Case 36609
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_CONTABIL_NAO_PREENCHIDO", gErr)
            
        Case 59221
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_JA_LANCADO", gErr)
        
        Case 110001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTABILIDADE_OBRIGATORIA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147858)

    End Select
    
    Exit Function

End Function

Private Function Contabil_Gera_Lancamento_Detalhe(colLancamento_Detalhe As Collection, lCliForn As Long, iFilialCliForn As Integer, lNumIntDocOrigem As Long, objLancamento_Cabecalho As ClassLancamento_Cabecalho, sModelo As String, sDocOrigem As String, iSubTipoTransacao As Integer) As Long
'Gera os lançamentos a partir do modelo sModelo e coloca os dados em colLancamento_Detalhe

Dim iIndice1 As Integer, iIndice2 As Integer
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim sConta As String
Dim sContaFormatada As String
Dim sCcl As String
Dim sCclFormatada As String
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim lErro As Long
Dim colLancamentos As New ClassColLancamentos
Dim objLancamento As ClassLancamentos
Dim colProdutos As New Collection
Dim colTotais As New Collection
Dim objProduto As New ClassProduto
Dim dTotal As Double

On Error GoTo Erro_Contabil_Gera_Lancamento_Detalhe

    lErro = CF("Gera_Lancamentos", sSiglaModulo, objTela, sModelo, colLancamentos, iSubTipoTransacao)
    If lErro <> SUCESSO Then gError 36597

    iIndice1 = 0
    
    For Each objLancamento In colLancamentos
    
        iIndice1 = iIndice1 + 1
    
        Set objLancamento_Detalhe = New ClassLancamento_Detalhe
        
        objLancamento_Detalhe.lCliForn = lCliForn
        objLancamento_Detalhe.iFilialCliForn = iFilialCliForn
        objLancamento_Detalhe.lNumIntDoc = lNumIntDocOrigem
        objLancamento_Detalhe.iTransacao = iCodTransacao
        objLancamento_Detalhe.iClienteFornecedor = iClienteFornecedor
        objLancamento_Detalhe.sDocOrigem = sDocOrigem
        objLancamento_Detalhe.sModelo = sModelo
                
        objLancamento_Detalhe.iSeqGrid = iIndice1
            
        sConta = objLancamento.sConta
           
        'se o campo de debito ou credito estiverem preenchidos ==> prossegue com a geracao
        If objLancamento.dDebito <> 0 Or objLancamento.dCredito <> 0 Then
        
            If Len(Trim(sConta)) = 0 Then gError 36598
                
            lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then gError 36599
                
            objLancamento_Detalhe.sConta = sContaFormatada
        
            If objLancamento.dDebito <> 0 And objLancamento.dCredito <> 0 Then gError 36601
                
            objLancamento_Detalhe.dValor = objLancamento.dCredito - objLancamento.dDebito
            
            iIndice2 = iIndice2 + 1
            
            objLancamento_Detalhe.iSeq = iIndice2
            
            'acumula os valores para posterior verificação do balanceamento do documento
            dTotal = dTotal + objLancamento_Detalhe.dValor
            
            objLancamento_Detalhe.iSeqContraPartida = objLancamento.iSeqContraPartida
            
            'Armazena Histórico e Ccl
            objLancamento_Detalhe.sHistorico = objLancamento.sHistorico
                
            'Se está usando Centro de Custo/Lucro, armazena-o
            If iGrid_Ccl_Col <> 999 Then
                    
                sCcl = objLancamento.sCcl
                
                lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
                If lErro <> SUCESSO Then gError 36602
                
                If iCclPreenchida = CCL_PREENCHIDA Then
                    objLancamento_Detalhe.sCcl = sCclFormatada
                Else
                    objLancamento_Detalhe.sCcl = ""
                End If
                    
            End If
                    
            'Armazena a indicação de aglutinação
            objLancamento_Detalhe.iAglutina = objLancamento.iAglutina
            
            objLancamento_Detalhe.sProduto = ""
                
            If giContabGerencial = 1 Then
                objLancamento_Detalhe.iGerencial = objLancamento.iGerencial
                If (objLancamento_Detalhe.iGerencial = 0 Or objLancamento_Detalhe.iGerencial = 2) And ((giFilialEmpresa > giFilialAuxiliar And giFilialAuxiliar > 0) Or (giFilialEmpresa < Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                    'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                    colLancamento_Detalhe.Add objLancamento_Detalhe
                End If
                    
                If (objLancamento_Detalhe.iGerencial = 1 Or objLancamento_Detalhe.iGerencial = 2) And ((giFilialEmpresa < giFilialAuxiliar And giFilialAuxiliar > 0) Or (giFilialEmpresa > Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                    colLancamento_Detalhe.Add objLancamento_Detalhe
                End If
                
            Else
                    
                'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                colLancamento_Detalhe.Add objLancamento_Detalhe
            End If
                
        End If
        
    Next
    
    'acerta o sequencial de contra partida pois alguns elementos do grid podem não ser armazenados por estarem zerados
    lErro = Acerta_Contra_Partida(colLancamento_Detalhe)
    If lErro <> SUCESSO Then gError 69044
    
    'verifica se o documento está balanceado
    If Round(dTotal, 2) <> 0 Then gError 55941
    
    Contabil_Gera_Lancamento_Detalhe = SUCESSO

    Exit Function

Erro_Contabil_Gera_Lancamento_Detalhe:
    
    Contabil_Gera_Lancamento_Detalhe = gErr

    Select Case gErr
    
        Case 36597
    
        Case 36598
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANC_CONTA_NAO_INFORMADA_DET", gErr, iCodTransacao, iSubTipoTransacao, sModelo, iIndice1, colLancamentos.Count, IIf(objLancamento.dDebito > 0, "-", "") & Format(objLancamento.dDebito + objLancamento.dCredito, "STANDARD"), sDocOrigem, lCliForn, iFilialCliForn, lNumIntDocOrigem)
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_LANC_CONTA_NAO_INFORMADA", gErr)
                
        Case 36599, 36602, 69044
        
        Case 36601
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANC_DEBCRE_PREENCHIDOS", gErr)
            
        Case 55941
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_NAO_BALANCEADO", gErr, sModelo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147859)
            
    End Select
    
    Exit Function

End Function

Function Contabil_DesfazContabilizacao1(iGrupoTransExc As Integer, lNumIntDocOrigem As Long, dtData As Date) As Long
'****************************** IMPORTANTE ************************************************
'especifica para tratar grupos de "transacoes" em que o lNumIntDocOrigem é o mesmo mas
'pode ter gerado lancamentos com codigos de transacao diferentes em diferentes filiais.
'Exemplo: baixa de parcela a pagar de uma filial efetuada por pagto de outra filial em uma Empresa
'em que as filiais tem autonomia contabil
'******************************************************************************************
'se os lancamentos identificados por lNumIntDocOrigem e uma transacao na tabela TransacaoCTB
'tenha GrupoTransExc igual ao passado como parametro, ainda nao tiverem
'sido lancados, exclui os lancamentos pendentes, caso já tiverem sido, faz o estorno contrario criando lancamentos inversos
'com data dtData

Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim colLancamento_Detalhe As New Collection, colLancamento_Detalhe2 As Collection
Dim objLanc As ClassLancamento_Detalhe
Dim objPeriodo As New ClassPeriodo, iFilialAnterior As Integer
Dim lDoc As Long, iIndice As Integer, iIndice2 As Integer
Dim objLancamento_Detalhe As ClassLancamento_Detalhe

On Error GoTo Erro_Contabil_DesfazContabilizacao1

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        'Consulta a tabela de Lançamentos com a chave NumIntDoc p/transacao em aiCodTransacao.
        'retorna lancamentos ordenados por filialempresa
        lErro = CF("Lancamento_Le_GrupoTransExc", iGrupoTransExc, lNumIntDocOrigem, colLancamento_Detalhe)
        If lErro <> SUCESSO And lErro <> 41988 Then Error 41995
     
        'Os lançamentos já foram contabilizados
        If lErro = SUCESSO Then
    
            iIndice = 1
            
            Do While iIndice < colLancamento_Detalhe.Count
        
                Set objLanc = colLancamento_Detalhe.Item(iIndice)
                
                lErro = CF("Periodo_Le1", dtData, objPeriodo, objLanc.iFilialEmpresa)
                If lErro <> SUCESSO Then Error 41996
                
                objLancamento_Cabecalho.iFilialEmpresa = objLanc.iFilialEmpresa
                objLancamento_Cabecalho.sOrigem = objLanc.sOrigem
                objLancamento_Cabecalho.iLote = 0
                objLancamento_Cabecalho.dtData = dtData
                objLancamento_Cabecalho.lNumIntDoc = lNumIntDocOrigem
                objLancamento_Cabecalho.iTransacao = objLanc.iTransacao
                
                'retorna um novo doc para o estorno
                lErro = CF("Voucher_Automatico_Trans", objLanc.iFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, objLancamento_Cabecalho.sOrigem, lDoc)
                If lErro <> SUCESSO Then Error 41997
                
                objLancamento_Cabecalho.lDoc = lDoc
                
                Set colLancamento_Detalhe2 = New Collection
                
                'inclui os lancamentos da filial numa colecao auxiliar
                For iIndice2 = iIndice To colLancamento_Detalhe.Count
                    
                    Set objLancamento_Detalhe = colLancamento_Detalhe.Item(iIndice2)
                    
                    'se encontrou lcto de outra filial interromper. Será processado no For mais externo a este
                    If objLancamento_Detalhe.iFilialEmpresa <> objLanc.iFilialEmpresa Then Exit For
                    
                    'inverte o sinal do lancamento
                    objLancamento_Detalhe.dValor = -objLancamento_Detalhe.dValor
                    
                    colLancamento_Detalhe2.Add objLancamento_Detalhe
                    
                Next
                
                'contabiliza os lançamentos imediatamente
                lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancamento_Detalhe2)
                If lErro <> SUCESSO Then Error 41998
                
                iIndice = iIndice + colLancamento_Detalhe2.Count
    
            Loop
            
        Else 'se o lancamento estiver "pendente"
                
            'Consulta a tabela de Lançamentos Pendentes com a chave NumIntDoc para as transacoes em aiCodTransacao
            lErro = CF("LanPendente_Le_GrupoTransExc", iGrupoTransExc, lNumIntDocOrigem, colLancamento_Detalhe)
            If lErro <> SUCESSO And lErro <> 41993 Then Error 41999
        
            'Se houverem lançamentos pendentes ==> exclui-los
            If lErro = SUCESSO Then
            
                iFilialAnterior = -1
                
                For Each objLanc In colLancamento_Detalhe
            
                    If iFilialAnterior <> objLanc.iFilialEmpresa Then
                    
                        iFilialAnterior = objLanc.iFilialEmpresa
                    
                        objLancamento_Cabecalho.iFilialEmpresa = objLanc.iFilialEmpresa
                        objLancamento_Cabecalho.sOrigem = objLanc.sOrigem
                        objLancamento_Cabecalho.iLote = objLanc.iLote
                        objLancamento_Cabecalho.dtData = objLanc.dtData
                        objLancamento_Cabecalho.lDoc = objLanc.lDoc
                    
                        lErro = CF("Lancamento_Exclui2", objLancamento_Cabecalho)
                        If lErro <> SUCESSO Then Error 32221
        
                    End If
                    
                Next
            
            End If

        End If

    End If
    
    Contabil_DesfazContabilizacao1 = SUCESSO
     
    Exit Function
    
Erro_Contabil_DesfazContabilizacao1:

    Contabil_DesfazContabilizacao1 = Err
     
    Select Case Err
          
        Case 41995 To 41999, 32221
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147860)
     
    End Select
     
    Exit Function

End Function

Function Contabil_Estorno(lNumIntDocOrigem As Long, dtData As Date, iFilialEmpresa As Integer, lDocParam As Long) As Long
'cria lançamentos que extornam os dados contábeis, se houverem, relativos a lNumIntDoc passado como parametro
'dtData indica a data com que o estorno será contabilizado e se iLote estiver preenchido, o lote que conterá os lançamentos deste estorno
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim colLancamento_Detalhe As New Collection
Dim colLancDet As New Collection
Dim objLanc As ClassLancamento_Detalhe
Dim objPeriodo As New ClassPeriodo
Dim lDoc As Long
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim iExercicio1 As Integer
Dim iPeriodoLan1 As Integer
Dim lDoc1 As Long
    
On Error GoTo Erro_Contabil_Estorno

    If iFilialEmpresa = 0 Then
        iFilialEmpresa = giFilialEmpresa
    End If
    
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        'le os lançamentos relativos ao documento passado como parametro ,coloca-os em colLancamento_Detalhe e atualiza o status para VOUCHER_ESTORNADO
        lErro = CF("Lancamento_Altera_Status", iFilialEmpresa, iCodTransacao, lNumIntDocOrigem, colLancamento_Detalhe, VOUCHER_ESTORNADO)
        If lErro <> SUCESSO And lErro <> 83876 Then gError 36846
    
        'Os lançamentos já foram contabilizados
        If lErro = SUCESSO Then
    
            lErro = CF("Periodo_Le", dtData, objPeriodo)
            If lErro <> SUCESSO Then gError 36845
            
            Set objLanc = colLancamento_Detalhe.Item(1)
            
            objLancamento_Cabecalho.iFilialEmpresa = objLanc.iFilialEmpresa
            objLancamento_Cabecalho.sOrigem = objLanc.sOrigem
            'os estornos sempre serão feitos sempre imediatamente pois senão poderia eventualmente deletar o lançamento de estorno no lote e o estornado ficaria marcado como tal.
            objLancamento_Cabecalho.iLote = 0
            objLancamento_Cabecalho.dtData = dtData
            objLancamento_Cabecalho.lNumIntDoc = lNumIntDocOrigem
            objLancamento_Cabecalho.iTransacao = iCodTransacao
            
            If lDocParam = 0 Then
            
                'retorna um novo doc para o estorno
                lErro = CF("Voucher_Automatico_Trans", objLancamento_Cabecalho.iFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, objLancamento_Cabecalho.sOrigem, lDoc)
                If lErro <> SUCESSO Then gError 36847
                
                lDocParam = lDoc
            
            Else
            
                lDoc = lDocParam
                
            End If
            
            objLancamento_Cabecalho.lDoc = lDoc
            
            lDoc1 = objLanc.lDoc
            iPeriodoLan1 = objLanc.iPeriodoLan
            iExercicio1 = objLanc.iExercicio
            
            'inverte os valores dos lançamentos a fim de extorná-los
            For Each objLancamento_Detalhe In colLancamento_Detalhe
                
                If objLancamento_Cabecalho.iFilialEmpresa <> objLancamento_Detalhe.iFilialEmpresa Or _
                   objLancamento_Cabecalho.sOrigem <> objLancamento_Detalhe.sOrigem Or _
                   iExercicio1 <> objLancamento_Detalhe.iExercicio Or _
                   iPeriodoLan1 <> objLancamento_Detalhe.iPeriodoLan Or _
                   lDoc1 <> objLancamento_Detalhe.lDoc Then
                   
                    'contabiliza os lançamentos imediatamente
                    lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancDet)
                    If lErro <> SUCESSO Then gError 36849
            
                    objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Detalhe.iFilialEmpresa
                    objLancamento_Cabecalho.sOrigem = objLancamento_Detalhe.sOrigem
                    iExercicio1 = objLancamento_Detalhe.iExercicio
                    iPeriodoLan1 = objLancamento_Detalhe.iPeriodoLan
                    lDoc1 = objLancamento_Detalhe.lDoc
                
                    'retorna um novo doc para o estorno
                    lErro = CF("Voucher_Automatico_Trans", objLancamento_Cabecalho.iFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, objLancamento_Cabecalho.sOrigem, lDoc)
                    If lErro <> SUCESSO Then gError 110043
                    
                    objLancamento_Cabecalho.lDoc = lDoc
                
                    Set colLancDet = New Collection
                
                End If
                
                '??? jones If dtData < objLancamento_Detalhe.dtData Then gError 92049
            
                objLancamento_Detalhe.dValor = -objLancamento_Detalhe.dValor
                objLancamento_Detalhe.dValorLivroAuxiliar = -objLancamento_Detalhe.dValorLivroAuxiliar
                objLancamento_Detalhe.iStatus = VOUCHER_ESTORNADOR
                
                colLancDet.Add objLancamento_Detalhe
                
            Next
            
            'contabiliza os lançamentos imediatamente
            lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancDet)
            If lErro <> SUCESSO Then gError 110044
    
        End If

    End If
    
    Contabil_Estorno = SUCESSO

    Exit Function

Erro_Contabil_Estorno:

    Contabil_Estorno = gErr

    Select Case gErr
    
        Case 36845, 36846, 36847, 36849, 110043, 110044
        
        Case 92049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_ESTORNO_DATA_MENOR", gErr, dtData, objLancamento_Detalhe.dtData, objLancamento_Detalhe.iFilialEmpresa, objLancamento_Detalhe.sOrigem, objLancamento_Detalhe.lDoc)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147861)

    End Select
    
    Exit Function

End Function

Sub Trata_SeqContraPartida(iLinhaExcluida As Integer)
'altera os indicadores de contra partida quando uma linha é excluida

Dim iLinha As Integer

    For iLinha = 1 To objGridInt.iLinhasExistentes
    
        If Len(objTela.Controls("CTBGridContabil").TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) > 0 Then
    
            If CInt(objTela.Controls("CTBGridContabil").TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) = iLinhaExcluida Then
                objTela.Controls("CTBGridContabil").TextMatrix(iLinha, iGrid_SeqContraPartida_Col) = ""
            ElseIf CInt(objTela.Controls("CTBGridContabil").TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) > iLinhaExcluida Then
                objTela.Controls("CTBGridContabil").TextMatrix(iLinha, iGrid_SeqContraPartida_Col) = CStr(CInt(objTela.Controls("CTBGridContabil").TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) - 1)
                
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147862)

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
        
            If objContraPartida.iSeqContraPartida = objLancamento_Detalhe.iSeqGrid Then
            
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_CONTRA_PARTIDA_VALOR", Err, objLancamento_Detalhe.iSeqGrid, Abs(objLancamento_Detalhe.dValor), Abs(objContraPartida.dValorContraPartida))
        
        Case 20617
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_CONTRA_PARTIDA_INEXISTENTE", Err, objContraPartida.iSeqContraPartida)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147863)

    End Select
    
    Exit Function

End Function

Private Sub Class_Terminate()

    'Código p/ descarregar referencias --> Jaques 13/1/2000
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    Set objTela = Nothing
    Set objGridInt = Nothing
    Set objEventoConta = Nothing
    Set objEventoCcl = Nothing
    Set objEventoHist = Nothing
    Set objContaButton = Nothing
    Set objCclButton = Nothing
    Set objHistButton = Nothing
    
    
End Sub


'???
Private Sub objEventoDoc_evSelecao(obj1 As Object)

    Call Contabil_objEventoDoc_evSelecao(obj1)

End Sub

'???
Private Sub objEventoLote_evSelecao(obj1 As Object)

    Call Contabil_objEventoLote_evSelecao(obj1)
    
End Sub

Private Function Acerta_Contra_Partida(colLancamento_Detalhe As Collection) As Long
'acerta os sequenciais de contra partida em função dos lançamentos zerados que não existirão.

Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim objLancamento_Detalhe1 As ClassLancamento_Detalhe
Dim lErro As Long

On Error GoTo Erro_Acerta_Contra_Partida

    For Each objLancamento_Detalhe In colLancamento_Detalhe
    
        If objLancamento_Detalhe.iSeqContraPartida <> 0 Then

            For Each objLancamento_Detalhe1 In colLancamento_Detalhe

                If objLancamento_Detalhe.iSeqContraPartida = objLancamento_Detalhe1.iSeqGrid Then
                
                    objLancamento_Detalhe.iSeqContraPartida = objLancamento_Detalhe1.iSeq
                    
                    Exit For

                End If
                
            Next
            
        End If

    Next
    
    Acerta_Contra_Partida = SUCESSO
    
    Exit Function

Erro_Acerta_Contra_Partida:

    Acerta_Contra_Partida = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147864)

    End Select
    
    Exit Function

End Function

Public Function Contabil_Inicializa_Contabilidade_Consulta(objTela1 As Object, objGridInt1 As AdmGrid, sSiglaModulo1 As String, Optional iSubTipoTransacao As Integer = 0) As Long

Dim lErro As Long
Dim objTransacaoCTB As New ClassTransacaoCTB

On Error GoTo Erro_Contabil_Inicializa_Contabilidade_Consulta

    iGeraModelo = 0

    Set objTela = objTela1
    
    iSubTipoTransacaoExt = iSubTipoTransacao

    Set objGridInt1 = New AdmGrid

    Set objGridInt = objGridInt1

    sSiglaModulo = sSiglaModulo1

    sNomeModulo = gcolModulo.Nome(sSiglaModulo)
    
    objTela.Controls("CTBOrigem").Caption = sNomeModulo

    'Tela em questão
    Set objGridInt.objForm = objTela1

'    If giContabGerencial = 1 Then
'
'        objTransacaoCTB.sSigla = sSiglaModulo
'        objTransacaoCTB.sTransacao = objTela1.Name
'        objTransacaoCTB.iSubTipo = iSubTipoTransacaoExt
'
'        lErro = CF("TransacaoCTB_Le", objTransacaoCTB)
'        If lErro <> SUCESSO Then gError 188765
'
'        objGridInt.iGerencial = objTransacaoCTB.iGerencial
'
'    End If
    
    'inicializa o grid de contabilizacao
    lErro = Contabil_Inicializa_GridContabil(True)
    If lErro <> SUCESSO Then gError 58950

    If objContab.giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        objTela.Controls("CTBCclLabel").Visible = True
        objTela.Controls("CTBCclDescricao").Visible = True
    End If

    Contabil_Inicializa_Contabilidade_Consulta = SUCESSO

    Exit Function

Erro_Contabil_Inicializa_Contabilidade_Consulta:

    Contabil_Inicializa_Contabilidade_Consulta = gErr

    Select Case gErr

        Case 58950, 188765

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147865)

    End Select

    Exit Function

End Function

Function Contabil_Traz_Doc_Tela_Consulta(lNumIntDoc As Long, sSiglaModulo1 As String, sTela As String) As Long
'traz os dados do voucher do banco de dados para a Tela de Consulta

Dim lErro As Long
Dim lErro1 As Long
Dim colLancamentos As New Collection
Dim objLanc As ClassLancamento_Detalhe
Dim iLinha As Integer
Dim sDescricao As String
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim iIndice As Integer
Dim dAcumulador As Double
Dim objTransacaoCTB As New ClassTransacaoCTB
Dim sModelo As String
Dim dtData As Date

On Error GoTo Erro_Contabil_Traz_Doc_Tela_Consulta

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        sSiglaModulo = sSiglaModulo1
    
        sNomeModulo = gcolModulo.Nome(sSiglaModulo)
    
        objTela.Controls("CTBOrigem").Caption = sNomeModulo
    
        objTransacaoCTB.sSigla = sSiglaModulo1
        objTransacaoCTB.sTransacao = sTela
    
        lErro = CF("TransacaoCTB_Le1", objTransacaoCTB)
        If lErro <> SUCESSO Then Error 58951
    
        iCodTransacao = objTransacaoCTB.iCodigo
    
        Call Contabil_Limpa_Contabilidade_Consulta
    
        'le os lançamentos pendentes contidos no documento em questão
        lErro = CF("LanPendente_Le_Doc3", giFilialEmpresa, iCodTransacao, lNumIntDoc, colLancamentos, sSiglaModulo, sModelo, dtData)
        If lErro <> SUCESSO And lErro <> 92019 Then Error 58952
    
        If giContabGerencial = 1 And ((giFilialEmpresa > giFilialAuxiliar And giFilialAuxiliar > 0) Or (giFilialEmpresa < Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
            'le os lançamentos pendentes contidos no documento em questão
            lErro1 = CF("LanPendente_Le_Doc3", giFilialEmpresa - giFilialAuxiliar, iCodTransacao, lNumIntDoc, colLancamentos, sSiglaModulo, sModelo, dtData)
            If lErro1 <> SUCESSO And lErro1 <> 92019 Then Error 58952
            
        Else
            lErro1 = 92019
            
        End If
    
    
'        For iIndice = 0 To objTela.Controls("CTBModelo").ListCount - 1
'            If objTela.Controls("CTBModelo").List(iIndice) = sModelo Then
'                objTela.Controls("CTBModelo").ListIndex = iIndice
'                Exit For
'            End If
'        Next
    
        'Inicializa Data
        If dtData <> 0 Then objTela.Controls("CTBDataContabil").Caption = Format(dtData, "dd/mm/yy")
    
        'se não encontrou o documento
        If lErro = 92019 And lErro1 = 92019 Then
    
            'le os lançamentos contidos no documento em questão
            lErro = CF("Lancamento_Le_Doc4", giFilialEmpresa, iCodTransacao, lNumIntDoc, colLancamentos, sSiglaModulo, sModelo, dtData)
            If lErro <> SUCESSO And lErro <> 92024 Then Error 58884
    
            If giContabGerencial = 1 And ((giFilialEmpresa > giFilialAuxiliar And giFilialAuxiliar > 0) Or (giFilialEmpresa < Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
    
                'le os lançamentos contidos no documento em questão
                lErro1 = CF("Lancamento_Le_Doc4", giFilialEmpresa - giFilialAuxiliar, iCodTransacao, lNumIntDoc, colLancamentos, sSiglaModulo, sModelo, dtData)
                If lErro1 <> SUCESSO And lErro1 <> 92024 Then Error 58884
    
            Else
                lErro1 = 92024
                
            End If
    
    
'            For iIndice = 0 To objTela.Controls("CTBModelo").ListCount - 1
'                If objTela.Controls("CTBModelo").List(iIndice) = sModelo Then
'                    objTela.Controls("CTBModelo").ListIndex = iIndice
'                    Exit For
'                End If
'            Next
    
            'Inicializa Data
            If dtData <> 0 Then objTela.Controls("CTBDataContabil").Caption = Format(dtData, "dd/mm/yy")
    
            'se não encontrou o documento
            If lErro = 92024 And lErro1 = 92024 Then Error 58885
    
        End If
    
        If lErro = SUCESSO Then
    
            Set objLanc = colLancamentos.Item(1)
    
            If objLanc.iLote <> 0 Then objTela.Controls("CTBLote").Caption = CStr(objLanc.iLote)
    
            objTela.Controls("CTBDocumento").Caption = CStr(objLanc.lDoc)
    
            'Coloca o periodo relativo a data na tela
            lErro = CF("Periodo_Le", objLanc.dtData, objPeriodo)
            If lErro <> SUCESSO Then Error 58886
    
            objTela.Controls("CTBPeriodo").Caption = objPeriodo.sNomeExterno
    
            'Coloca o exercicio na tela
            lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
            If lErro <> SUCESSO And lErro <> 10083 Then Error 58893
    
            'Se o exercicio não está cadastrado
            If lErro = 10083 Then Error 58894
    
            objTela.Controls("CTBExercicio").Caption = objExercicio.sNomeExterno
    
            'move os dados para a tela
            lErro = Contabil_Traz_Doc_Tela1(colLancamentos)
            If lErro <> SUCESSO Then Error 58895
    
            dAcumulador = Contabil_GridColuna_Soma(iGrid_Credito_Col)
            objTela.Controls("CTBTotalCredito").Caption = Format(dAcumulador, "Standard")
    
            dAcumulador = Contabil_GridColuna_Soma(iGrid_Debito_Col)
            objTela.Controls("CTBTotalDebito").Caption = Format(dAcumulador, "Standard")
    
        End If

    End If
    
    Contabil_Traz_Doc_Tela_Consulta = SUCESSO

    Exit Function

Erro_Contabil_Traz_Doc_Tela_Consulta:

    Contabil_Traz_Doc_Tela_Consulta = Err

    Select Case Err

        Case 58884, 58885, 58886, 58893, 58895, 58951, 58952

        Case 58894
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147866)

    End Select

    Exit Function

End Function

Function Contabil_Limpa_Contabilidade_Consulta() As Long
'Limpa a Contabilidade nas Telas de Consulta

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        Call Grid_Limpa(objGridInt)
        objTela.Controls("CTBTotalDebito").Caption = ""
        objTela.Controls("CTBTotalCredito").Caption = ""
        objTela.Controls("CTBContaDescricao").Caption = ""
        objTela.Controls("CTBCclDescricao").Caption = ""
        objTela.Controls("CTBDocumento").Caption = ""
        objTela.Controls("CTBLote").Caption = ""
        objTela.Controls("CTBDataContabil").Caption = "  /  /  "
        objTela.Controls("CTBPeriodo").Caption = ""
        objTela.Controls("CTBExercicio").Caption = ""

    End If

    Contabil_Limpa_Contabilidade_Consulta = SUCESSO

End Function

Function Contabil_BotaoImprimir_Consulta_Click() As Long

Dim lErro As Long, objRelTela As New ClassRelTela, iIndice1 As Integer
Dim colTemp As Collection, objLancamento As New ClassLancamento_Detalhe
Dim colLancamento_Detalhe As New Collection
Dim sOrigem As String, sExercicio As String, sPeriodo As String
Dim lDoc As Long, dtData As Date, lLote As Long
Dim iFilialEmpresaSalva As Integer
Dim iIndice As Integer

On Error GoTo Erro_Contabil_BotaoImprimir_Consulta_Click

    iFilialEmpresaSalva = giFilialEmpresa

    lErro = objRelTela.Iniciar("REL_LANC_CTB")
    If lErro <> SUCESSO Then gError 60013

    'obter dados comuns a todas as linhas do grid
    sOrigem = objContab.gobjColOrigem.Origem(objTela.Controls("CTBOrigem").Caption)
    sExercicio = objTela.Controls("CTBExercicio").Caption
    sPeriodo = objTela.Controls("CTBPeriodo").Caption

    lDoc = StrParaLong(objTela.Controls("CTBDocumento").Caption)
    dtData = StrParaDate(objTela.Controls("CTBDataContabil").Caption)
    lLote = StrParaLong(objTela.Controls("CTBLote").Caption)

    'Preenche Objeto Lançamento_Detalhe
    lErro = Contabil_Grid_Lancamento_Detalhe(colLancamento_Detalhe, 0, 0, 0, -1, "", True)
    If lErro <> SUCESSO Then gError 60014

    If giContabGerencial = 1 Then
    
        If ((giFilialEmpresa > giFilialAuxiliar And giFilialAuxiliar > 0) Or (giFilialEmpresa < Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
        
            giFilialEmpresa = giFilialEmpresa - giFilialAuxiliar
            
            'Preenche Objeto Lançamento_Detalhe
            lErro = Contabil_Grid_Lancamento_Detalhe(colLancamento_Detalhe, 0, 0, 0, -1, "", True)
            If lErro <> SUCESSO Then gError 188072
            
            giFilialEmpresa = iFilialEmpresaSalva
            
            For iIndice = colLancamento_Detalhe.Count To 1 Step -1
                
                For iIndice1 = iIndice - 1 To 1 Step -1
                
                    If colLancamento_Detalhe(iIndice).iSeqGrid = colLancamento_Detalhe(iIndice1).iSeqGrid Then
                        colLancamento_Detalhe.Remove (iIndice)
                        Exit For
                    End If
                    
                Next
            
            Next
                        
        End If
        
    End If


    For iIndice1 = 1 To colLancamento_Detalhe.Count

        Set objLancamento = colLancamento_Detalhe.Item(iIndice1)

        Set colTemp = New Collection

        'incluir os valores na mesma ordem da tabela RelTelaCampos no dicdados

        Call colTemp.Add(sOrigem)
        Call colTemp.Add(sExercicio)
        Call colTemp.Add(sPeriodo)
        Call colTemp.Add(lDoc)
        Call colTemp.Add(iIndice1)
        Call colTemp.Add(lLote)
        Call colTemp.Add(dtData)
        Call colTemp.Add(objLancamento.sConta)
        Call colTemp.Add(objLancamento.sCcl)
        Call colTemp.Add(objLancamento.sHistorico)
        Call colTemp.Add(objLancamento.dValor)
        Call colTemp.Add(objLancamento.iSeqContraPartida)

        lErro = objRelTela.IncluirRegistro(colTemp)
        If lErro <> SUCESSO Then gError 60015

    Next

    lErro = objRelTela.ExecutarRel
    If lErro <> SUCESSO Then gError 60016

    Contabil_BotaoImprimir_Consulta_Click = SUCESSO

    Exit Function

Erro_Contabil_BotaoImprimir_Consulta_Click:

    Contabil_BotaoImprimir_Consulta_Click = gErr

    giFilialEmpresa = iFilialEmpresaSalva

    Select Case gErr

        Case 60013, 60014, 60015, 60016, 188072

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147867)

    End Select

    Exit Function

End Function

Function Contabil_GridContabil_Consulta_Click() As Long
'Para colocar a descricao da Conta e CCL na Tela, para Telas de Consulta

Dim lErro As Long
Dim sConta As String
Dim sCcl As String

On Error GoTo Erro_Contabil_GridContabil_Consulta_Click

    'Verifica se a Linha é maior que zero
    If objTela.Controls("CTBGridContabil").Row > 0 Then

        'se estiver clicando em uma celula de Conta
        If objTela.Controls("CTBGridContabil").col = GRID_COLUNA_CONTA Then

            sConta = objTela.Controls("CTBGridContabil").TextMatrix(objTela.Controls("CTBGridContabil").Row, objTela.Controls("CTBGridContabil").col)
    
            'Se a conta esta Preenchida
            If Len(sConta) > 0 Then
    
                lErro = Contabil_Conta_Exibe_Descricao(sConta)
                If lErro <> SUCESSO Then Error 60017
    
            Else
    
                objTela.Controls("CTBContaDescricao") = ""
    
            End If

        End If
        
        'Se existe CCL no Grid e estiver clicando em uma celula de centro de custo
        If objContab.giSetupUsoCcl = CCL_USA_EXTRACONTABIL And objTela.Controls("CTBGridContabil").col = GRID_COLUNA_CCL Then

            sCcl = objTela.Controls("CTBGridContabil").TextMatrix(objTela.Controls("CTBGridContabil").Row, objTela.Controls("CTBGridContabil").col)

            'Coloca descricao de Ccl no panel
            If Len(sCcl) > 0 Then

                lErro = Contabil_Ccl_Exibe_Descricao(sCcl)
                If lErro <> SUCESSO Then Error 60018

            Else

                objTela.Controls("CTBCclDescricao") = ""

            End If

        End If

    End If

    Contabil_GridContabil_Consulta_Click = SUCESSO

    Exit Function

Erro_Contabil_GridContabil_Consulta_Click:

    Contabil_GridContabil_Consulta_Click = Err

    Select Case Err

        Case 60017, 60018

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147868)

    End Select

    Exit Function

End Function

Function Contabil_Trata_Estorno_Versao_Light(ByVal iEstorno As Integer) As Long
'se for um estorno, gera o numero do documento de estorno

Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim lDoc As Long
Dim iAchou As Integer
Dim iIndice As Integer

On Error GoTo Erro_Contabil_Trata_Estorno_Versao_Light

    'se algum elemento da transacao está sendo estornado
    If iEstorno = 1 Then

        objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
        objLancamento_Cabecalho.sOrigem = sSiglaModulo
        objLancamento_Cabecalho.iExercicio = iExercicio
        objLancamento_Cabecalho.iPeriodoLan = iPeriodo
        If Len(objTela.Controls("CTBDocumento").Text) > 0 Then objLancamento_Cabecalho.lDoc = CLng(objTela.Controls("CTBDocumento").Text)
    
        lErro = CF("Lancamento_Le_Doc", objLancamento_Cabecalho)
        If lErro <> SUCESSO And lErro <> 5726 Then Error 60875

        'se o documento em questão já está cadastrado, pega um novo numero
        If lErro = SUCESSO Then

            'cria o lDoc do documento de estorno
            lErro = CF("Voucher_Automatico", giFilialEmpresa, iExercicio, iPeriodo, sSiglaModulo, lDoc)
            If lErro <> SUCESSO Then Error 60864

            objTela.Controls("CTBDocumento").Text = CStr(lDoc)
        End If

        'se o modelo padrão não estiver selecionado ==> selecioná-lo
        If objTela.Controls("CTBModelo").ListIndex = -1 Then

            'pesquisa na combobox Modelo o modelo padrão
            For iIndice = 0 To objTela.Controls("CTBModelo").ListCount - 1
                If objTela.Controls("CTBModelo").ItemData(iIndice) = PADRAOCONTAB_MODELO_PADRAO Then
                    iAchou = 1
                    Exit For
                End If
            Next

            If iAchou = 1 Then objTela.Controls("CTBModelo").ListIndex = iIndice

        End If

        objTela.Controls("CTBLancAutomatico").Value = MARCADO

    End If

    Contabil_Trata_Estorno_Versao_Light = SUCESSO

    Exit Function

Erro_Contabil_Trata_Estorno_Versao_Light:

    Contabil_Trata_Estorno_Versao_Light = Err

    Select Case Err

        Case 60864, 60875

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147869)

    End Select

    Exit Function

End Function

Function Contabil_Gera_Cabecalho_Automatico() As Long
'preenche valores default p/data contabil, exercicio

Dim lErro As Long
Dim iGeraLoteAutomatico As Integer
Dim objLote As New ClassLote
Dim iLote As Integer

On Error GoTo Erro_Contabil_Gera_Cabecalho_Automatico

    'verifica se deve gerar numeração de lote automaticamente
    lErro = Verifica_Gera_Lote_Automatico(iGeraLoteAutomatico)
    If lErro <> SUCESSO Then Error 59402

    'traz a data/periodo/exercicio atuais
    lErro = Contabil_Traz_Cabecalho_Tela()
    If lErro <> SUCESSO Then Error 59403

    'se é para gerar o lote automatico
    If iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO Then

        'verifica se o usuario já tem um numero de lote associado
        lErro = CF("Lote_Gera_Automatico", giFilialEmpresa, gsUsuario, sSiglaModulo, iExercicio, iPeriodo, iLote)
        If lErro <> SUCESSO And lErro <> 60885 And lErro <> 60890 Then Error 59404

        'se não há lote pendente disponivel para a origem e periodo em questão ==> erro
        If lErro <> SUCESSO Then
        
            Call Rotina_Aviso(vbOKOnly, "ERRO_NAO_HA_LOTE_PENDENTE1", sSiglaModulo, iPeriodo, iExercicio)

        Else
        
            'se o lote foi encontrado
            objTela.Controls("CTBLote").Text = CStr(iLote)
        
        End If
        
    End If

    Contabil_Gera_Cabecalho_Automatico = SUCESSO
     
    Exit Function
    
Erro_Contabil_Gera_Cabecalho_Automatico:

    Contabil_Gera_Cabecalho_Automatico = Err
     
    Select Case Err
          
        Case 59402, 59403, 59404
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147870)
     
    End Select
     
    Exit Function

End Function

Function Contabil_LancAutomatico_Click() As Long

On Error GoTo Erro_Contabil_LancAutomatico_Click

    If objTela.Controls("CTBLancAutomatico").Value = MARCADO And objTela.Controls("CTBModelo").ListIndex = -1 Then
        objTela.Controls("CTBLancAutomatico").Value = DESMARCADO
        Error 60906
    End If
   
    Contabil_LancAutomatico_Click = SUCESSO

    Exit Function
    
Erro_Contabil_LancAutomatico_Click:

    Contabil_LancAutomatico_Click = Err
     
    Select Case Err
          
        Case 60906
            Call Rotina_Erro(vbOKOnly, "ERRO_RECALCULO_AUTOMATICO_SEM_MODELO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147871)
     
    End Select
     
    Exit Function

End Function

Private Sub Historico_Troca_Num_Doc(sHistorico As String, lNumDoc As Long)

Dim iPosicao As Integer

    iPosicao = InStr(sHistorico, NUMERO_NOTA_FISCAL)
    
    Do While iPosicao <> 0
    
        sHistorico = left(sHistorico, iPosicao - 1) & CStr(lNumDoc) & Mid(sHistorico, iPosicao + 4)
    
        iPosicao = iPosicao + 4
        
        iPosicao = InStr(sHistorico, NUMERO_NOTA_FISCAL)
    
    Loop

    iPosicao = InStr(sHistorico, HIST_CTB_NUMERO_FAT)
    
    Do While iPosicao <> 0
    
        sHistorico = left(sHistorico, iPosicao - 1) & CStr(lNumDoc) & Mid(sHistorico, iPosicao + 4)
    
        iPosicao = iPosicao + 4
        
        iPosicao = InStr(sHistorico, HIST_CTB_NUMERO_FAT)
    
    Loop

End Sub

Private Function Contabil_Carga_Modelos() As Long

Dim colPadraoContab As New Collection
Dim objPadraoContab As ClassPadraoContab
Dim objTransacaoCTB As New ClassTransacaoCTB
Dim iCodTransacao1 As Integer
Dim lErro As Long
Dim iAchou As Integer
Dim iIndice As Integer

On Error GoTo Erro_Contabil_Carga_Modelos

    If Not objTela Is Nothing Then

        lErro = Contabil_Inicializa_Contabilidade2()
        If lErro <> SUCESSO Then gError 92138

        objTransacaoCTB.sSigla = sSiglaModulo
        objTransacaoCTB.sTransacao = objTela.Name
        objTransacaoCTB.iSubTipo = iSubTipoTransacaoExt
        
        lErro = CF("TransacaoCTB_Le", objTransacaoCTB)
        If lErro <> SUCESSO Then gError 71777
    
        iCodTransacao = objTransacaoCTB.iCodigo
        iClienteFornecedor = objTransacaoCTB.iCliFornecedor
        
     
        iObrigatoriedade = objTransacaoCTB.iObrigatoriedade
    
        iCodTransacao1 = iCodTransacao
    
        'leitura dos modelos de contabilizacao referentes ao modulo/transacao em questao
        lErro = CF("PadraoContab_Le_Todos", sSiglaModulo, iCodTransacao1, colPadraoContab, iSubTipoTransacaoExt)
        If lErro <> SUCESSO Then gError 71778
    
        objTela.Controls("CTBModelo").Clear
    
        For Each objPadraoContab In colPadraoContab
    
            objTela.Controls("CTBModelo").AddItem objPadraoContab.sModelo
            objTela.Controls("CTBModelo").ItemData(objTela.Controls("CTBModelo").NewIndex) = objPadraoContab.iPadrao
    
        Next
    
        'se o recalculo estiver marcado procura selecionar o modelo padrão se houver
        If objTela.Controls("CTBLancAutomatico").Value = MARCADO Then
        
            Call Contabil_Limpa_GridContabil1
        
            'pesquisa na combobox Modelo o modelo padrão
            For iIndice = 0 To objTela.Controls("CTBModelo").ListCount - 1
                If objTela.Controls("CTBModelo").ItemData(iIndice) = PADRAOCONTAB_MODELO_PADRAO Then
                    iAchou = 1
                    Exit For
                End If
            Next
    
            If iAchou = 1 Then
                objTela.Controls("CTBModelo").ListIndex = iIndice
                objTela.Controls("CTBLancAutomatico").Value = MARCADO
            End If
        
        End If

    End If

    Contabil_Carga_Modelos = SUCESSO
    
    Exit Function

Erro_Contabil_Carga_Modelos:

    Contabil_Carga_Modelos = gErr
    
    Select Case gErr
    
        Case 71777, 71778, 92138

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147872)

    End Select
    
    Exit Function

End Function

Public Function Contabil_Processa_Alteracao_Subtipo(iSubTipoTransacao As Integer) As Long
'usado para processar a alteração do subtipo em uma tela

Dim lErro As Long

On Error GoTo Erro_Contabil_Processa_Alteracao_Subtipo

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        iSubTipoTransacaoExt = iSubTipoTransacao
    
        lErro = Contabil_Carga_Modelos()
        If lErro <> SUCESSO Then gError 71779

    End If

    Contabil_Processa_Alteracao_Subtipo = SUCESSO
    
    Exit Function
    
Erro_Contabil_Processa_Alteracao_Subtipo:

    Contabil_Processa_Alteracao_Subtipo = Err
    
    Select Case Err
    
        Case 71779
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147873)

    End Select
    
    Exit Function

End Function

Function Contabil_Testa_Data(dtData As Date) As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim dtDataContabil As Date

On Error GoTo Erro_Contabil_Testa_Data

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        If Len(objTela.Controls("CTBDataContabil").ClipText) = 0 Then
            dtDataContabil = gdtDataAtual
        Else
            dtDataContabil = CDate(objTela.Controls("CTBDataContabil").Text)
        End If
    
        If dtDataContabil <> dtData Then
    
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DATA_CONTABIL_DIFERE_TRANSACAO", dtDataContabil, dtData)
    
                If vbMsgRes = vbNo Then gError 92027
    
        End If

    End If

    Contabil_Testa_Data = SUCESSO

    Exit Function

Erro_Contabil_Testa_Data:

    Contabil_Testa_Data = gErr

    Select Case gErr

        Case 92027

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147874)

    End Select

    Exit Function

End Function

Public Function Contabil_Inicializa_Contabilidade4(ByVal iCodTransacaoParam As Integer, ByVal sSiglaModuloParam As String) As Long
    iCodTransacao = iCodTransacaoParam
    sSiglaModulo = sSiglaModuloParam
End Function

Private Function Contabil_Exclui1(lNumIntDocOrigem As Long, Optional iSemEstornar1 As Integer = 0, Optional dtData As Date, Optional ByVal sOrigem As String, Optional ByVal dtDataEstorno As Date = DATA_NULA, Optional ByVal iFilialEmpresa As Integer = 0, Optional lDocParam As Long = 0) As Long
'exclui os dados contábeis, se houverem, relativos a lNumIntDoc passado como parametro
'o parametro iSemEstornar indica se deve (ou nao) ser estornada contabilizacao já lancada
'se iSemEstornar = 2 significa que só deve estornar se o documento da tela for nulo
'se iSemEstornar = 3 indica que os lançamento já contabilizados devem ser excluidos
'devolve a data no caso do lançamento já tiver sido contabilizado. Esta data servirá para verificar se é um caso de alteração de data.
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim colLancamentos As New Collection
Dim objLanc As ClassLancamento_Detalhe
Dim colLancamentos1 As New Collection, iSemEstornar As Integer
    
On Error GoTo Erro_Contabil_Exclui1

    If iFilialEmpresa = 0 Then
        iFilialEmpresa = giFilialEmpresa
    End If
    
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
        If Len(Trim(sOrigem)) = 0 Then
            sOrigem = sSiglaModulo
        End If
    
        If GL_objContabInt.mvargiCtbLctoExcluiEstorno = CONTABIL_EXCLUI_EFETIVO Then
            iSemEstornar = CONTABIL_EXCLUI_EFETIVO
        Else
            iSemEstornar = iSemEstornar1
        End If

        If iSemEstornar = CONTABIL_EXCLUI_EFETIVO Then
        
            'Consulta a tabela de Lançamentos com a chave FilialEmpresa, Transacao, NumIntDoc. Se estiver cadastrado ==> vai estornar
            lErro = CF("Lancamento_Le_Doc3", iFilialEmpresa, iCodTransacao, lNumIntDocOrigem, colLancamentos)
            If lErro <> SUCESSO And lErro <> 89100 Then gError 89103
        
        Else
        
            'Consulta a tabela de Lançamentos com a chave FilialEmpresa, Transacao, NumIntDoc. Se estiver cadastrado ==> vai estornar
            lErro = CF("Lancamento_Le_Doc2", iFilialEmpresa, iCodTransacao, lNumIntDocOrigem, colLancamentos)
            If lErro <> SUCESSO And lErro <> 36295 Then gError 36388
    
        End If
        
        'Os lançamentos já foram contabilizados
        If lErro = SUCESSO Then
            
            dtData = colLancamentos(1).dtData
            
            If iSemEstornar = CONTABIL_EXCLUI_SEM_ESTORNO Then gError 59220
            
            'iSemEstornar = CONTABIL_EXCLUI_ESTORNO_DOC_ZERADO significa que pode estornar desde que o numero do documento na tela esteja nulo
            If iSemEstornar = CONTABIL_EXCLUI_ESTORNO_DOC_ZERADO Then
            
                If Len(objTela.Controls("CTBDocumento").ClipText) <> 0 Then gError 59220
            
            End If
            
            'essa funcao só serve para controlar a exclusao pelo dicionario de dados
            lErro = CF("Lancamento_Exclui_Cust")
            If lErro <> SUCESSO Then gError 204868

            
            'se iSemEstornar = CONTABIL_EXCLUI_EFETIVO ==> excluir fisicamente os lançamentos associados a transacao
            If iSemEstornar = CONTABIL_EXCLUI_EFETIVO Then
                
                lErro = CF("Lancamento_Exclui_30", iFilialEmpresa, iCodTransacao, lNumIntDocOrigem, colLancamentos1)
                If lErro <> SUCESSO Then gError 83885
                
            Else
                
                If dtDataEstorno = DATA_NULA Then
                    dtDataEstorno = dtData
                End If


                'criar a contabilizacao de estorno
                lErro = Contabil_Estorno(lNumIntDocOrigem, dtDataEstorno, iFilialEmpresa, lDocParam)
                If lErro <> SUCESSO Then gError 36389
            
            End If
        End If
        
        Set colLancamentos = New Collection
        
        'Exclui os Lancamentos pendentes se existirem
        'Consulta a tabela de Lançamentos Pendentes com a chave FilialEmpresa, Transacao, NumIntDoc
        'lErro = CF("LanPendente_Le_Doc2", iFilialEmpresa, iCodTransacao, lNumIntDocOrigem, colLancamentos, sSiglaModulo)
        lErro = CF("LanPendente_Le_Doc2", iFilialEmpresa, iCodTransacao, lNumIntDocOrigem, colLancamentos, sOrigem)
        If lErro <> SUCESSO And lErro <> 36290 Then gError 36390
    
        'Se houverem lançamentos pendentes ==> exclui-los
        If lErro = SUCESSO Then
        
            'essa funcao só serve para controlar a exclusao pelo dicionario de dados
            lErro = CF("Lancamento_Exclui_Cust")
            If lErro <> SUCESSO Then gError 204868

        
            Set objLanc = colLancamentos.Item(1)
        
            objLancamento_Cabecalho.iFilialEmpresa = objLanc.iFilialEmpresa
            objLancamento_Cabecalho.sOrigem = objLanc.sOrigem
            objLancamento_Cabecalho.iLote = objLanc.iLote
            objLancamento_Cabecalho.dtData = objLanc.dtData
            objLancamento_Cabecalho.lDoc = objLanc.lDoc
        
            lErro = CF("Lancamento_Exclui0", objLancamento_Cabecalho)
            If lErro <> SUCESSO Then gError 36393
        
        End If
        
        Set colLancamentos = New Collection
        
        'Exclui os Custos se existirem em Lanpedente
        'Consulta a tabela de Lançamentos Pendentes com a chave FilialEmpresa, Transacao, NumIntDoc
        lErro = CF("LanPendente_Le_Doc2", iFilialEmpresa, iCodTransacao, lNumIntDocOrigem, colLancamentos, sOrigemCusto)
        If lErro <> SUCESSO And lErro <> 36290 Then gError 64470
    
        'Se houverem Custos pendentes ==> exclui-los
        If lErro = SUCESSO Then
        
            Set objLanc = colLancamentos.Item(1)
        
            objLancamento_Cabecalho.iFilialEmpresa = objLanc.iFilialEmpresa
            objLancamento_Cabecalho.sOrigem = objLanc.sOrigem
            objLancamento_Cabecalho.iLote = objLanc.iLote
            objLancamento_Cabecalho.dtData = objLanc.dtData
            objLancamento_Cabecalho.lDoc = objLanc.lDoc
        
            lErro = CF("Lancamento_Exclui0", objLancamento_Cabecalho)
            If lErro <> SUCESSO Then gError 64471
        
        End If
    
    End If
    
    Contabil_Exclui1 = SUCESSO

    Exit Function

Erro_Contabil_Exclui1:

    Contabil_Exclui1 = gErr

    Select Case gErr
    
        Case 36388, 36389, 36390, 36393, 59220, 64470, 64471, 83885, 89103, 204868

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147846)

    End Select
    
    Exit Function

End Function

Function Contabil_Gerencial_Click() As Long

    objTela.iAlterado = REGISTRO_ALTERADO
    iContabilidade_Alterada = CONTABILIDADE_ALTERADA
   objTela.Controls("CTBLancAutomatico").Value = DESMARCADO
    Contabil_Gerencial_Click = SUCESSO

End Function

Function Contabil_Gerencial_GotFocus() As Long
'trata o evento gotfocus associado ao campo gerencial

    Call Grid_Campo_Recebe_Foco(objGridInt)

    Contabil_Gerencial_GotFocus = SUCESSO

End Function

Function Contabil_Gerencial_KeyPress(iKeyAscii As Integer) As Long
'trata o evento keypress associado ao campo Gerencial

    Call Grid_Trata_Tecla_Campo_Contabil(iKeyAscii, objGridInt)

    Contabil_Gerencial_KeyPress = SUCESSO

End Function

Function Contabil_Gerencial_Validate(Cancel As Boolean) As Long
'trata o evento validate associado ao campo Gerencial

Dim lErro As Long

    Set objGridInt.objControle = objTela.Controls("CTBGerencial")
    lErro = Grid_Campo_Libera_Foco_Contabil(objGridInt)
    If lErro <> SUCESSO Then Cancel = True
    
    Contabil_Gerencial_Validate = SUCESSO

End Function

Private Sub objContaButton_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(objTela.Controls("CTBConta").Text)) > 0 Then objPlanoConta.sConta = objTela.Controls("CTBConta").Text
    
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
    
    If objTela.Controls("CTBGridContabil").col = iGrid_Conta_Col Then

        Set objPlanoConta = obj1
        
        sConta = objPlanoConta.sConta
        
        'le a conta
        lErro = CF("PlanoConta_Le_Conta1", sConta, objPlanoConta)
        If lErro <> SUCESSO And lErro <> 6030 Then gError 197898
        
        If objPlanoConta.iAtivo <> CONTA_ATIVA Then gError 197902
        
        If objPlanoConta.iTipoConta <> CONTA_ANALITICA Then gError 197903
        
        sContaEnxuta = String(STRING_CONTA, 0)

        lErro = objContab.Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 197899

        objTela.Controls("CTBConta").PromptInclude = False
        objTela.Controls("CTBConta").Text = sContaEnxuta
        objTela.Controls("CTBConta").PromptInclude = True

        objTela.Controls("CTBGridContabil").TextMatrix(objTela.Controls("CTBGridContabil").Row, objTela.Controls("CTBGridContabil").col) = objTela.Controls("CTBConta").Text

        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

        objTela.Controls("CTBContaDescricao").Caption = objPlanoConta.sDescConta

        'Se a Conta possui um Histórico Padrão "default" coloca na tela
        If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Historico_Col))) = 0 And objPlanoConta.iHistPadrao <> 0 Then
                        
            objHistPadrao.iHistPadrao = objPlanoConta.iHistPadrao
                        
            'le os dados do historico
            lErro = CF("HistPadrao_Le", objHistPadrao)
            If lErro <> SUCESSO And lErro <> 5446 Then gError 197900
                                    
            If lErro = SUCESSO Then
            
                objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Historico_Col) = objHistPadrao.sDescHistPadrao
                
            End If
            
        End If

    End If
    
    objTela.Show
    
    Exit Sub
    
Erro_objEventoConta_evSelecao:

    Select Case gErr
    
        Case 197898, 197900
    
        Case 197899
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
    
        Case 197902
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", gErr, sConta)
        
        Case 197903
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", gErr, sConta)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197901)
        
    End Select

    Exit Sub

End Sub

Private Sub objCclButton_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

    If Len(Trim(objTela.Controls("CTBCcl").Text)) > 0 Then objCcl.sCcl = objTela.Controls("CTBCcl").Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
    
Dim sCcl As String
Dim sCclEnxuta As String
Dim lErro As Long
Dim sCaracterInicial As String
Dim objCcl As ClassCcl
    
On Error GoTo Erro_objEventoCcl_evSelecao
    
    If objTela.Controls("CTBGridContabil").col = iGrid_Ccl_Col Then

        Set objCcl = obj1

        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 197904

        If objCcl.iTipoCcl <> CCL_ANALITICA Then gError 197905
        
        If objCcl.iAtivo = 0 Then gError 197906

        sCclEnxuta = String(STRING_CCL, 0)

        'volta mascarado apenas os caracteres preenchidos
        lErro = objContab.Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclEnxuta)
        If lErro <> SUCESSO Then gError 197907

        objTela.Controls("CTBCcl").PromptInclude = False
        objTela.Controls("CTBCcl").Text = sCclEnxuta
        objTela.Controls("CTBCcl").PromptInclude = True

        objTela.Controls("CTBGridContabil").TextMatrix(objTela.Controls("CTBGridContabil").Row, objTela.Controls("CTBGridContabil").col) = objTela.Controls("CTBCcl").Text

        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
        objTela.Controls("CTBCclDescricao").Caption = objCcl.sDescCcl

    End If

    objTela.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 197904

        Case 197905
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_ANALITICA1", gErr, objCcl.sCcl)
  
        Case 197906
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_INATIVO", gErr, objCcl.sCcl)

        Case 197907
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197908)

    End Select

    Exit Sub

End Sub

Private Sub objHistButton_Click()

Dim colSelecao As Collection
Dim objHistPadrao As New ClassHistPadrao

    Call Chama_Tela("HistPadraoLista", colSelecao, objHistPadrao, objEventoHist)

End Sub

Private Sub objEventoHist_evSelecao(obj1 As Object)


Dim objHistPadrao As ClassHistPadrao

On Error GoTo Erro_objEventoHist_evSelecao

    If objTela.Controls("CTBGridContabil").col = iGrid_Historico_Col Then

        Set objHistPadrao = obj1

        objTela.Controls("CTBGridContabil").TextMatrix(objTela.Controls("CTBGridContabil").Row, objTela.Controls("CTBGridContabil").col) = objHistPadrao.sDescHistPadrao

        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    objTela.Show
    
    Exit Sub

Erro_objEventoHist_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197909)

    End Select

    Exit Sub

End Sub

