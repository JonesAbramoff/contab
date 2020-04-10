VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTBaixaRecTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Const NUM_MAX_PARCELAS_BAIXA = 300

Public Property Get NumVou(ByVal objCT As Object) As Object
     Set NumVou = objCT.objUserControl.Controls("NumVou")
End Property

Public Property Get NumCCred(ByVal objCT As Object) As Object
     Set NumCCred = objCT.objUserControl.Controls("NumCCred")
End Property

Public Property Get ClienteRA(ByVal objCT As Object) As Object
     Set ClienteRA = objCT.objUserControl.Controls("ClienteRA")
End Property

Public Property Get FilialClienteRA(ByVal objCT As Object) As Object
     Set FilialClienteRA = objCT.objUserControl.Controls("FilialClienteRA")
End Property

Public Property Get ValorDe(ByVal objCT As Object) As Object
     Set ValorDe = objCT.objUserControl.Controls("ValorDe")
End Property

Public Property Get ValorAte(ByVal objCT As Object) As Object
     Set ValorAte = objCT.objUserControl.Controls("ValorAte")
End Property

Public Property Get DataRADe(ByVal objCT As Object) As Object
     Set DataRADe = objCT.objUserControl.Controls("DataRADe")
End Property

Public Property Get DataRAAte(ByVal objCT As Object) As Object
     Set DataRAAte = objCT.objUserControl.Controls("DataRAAte")
End Property

Public Property Get UpDownRADe(ByVal objCT As Object) As Object
     Set UpDownRADe = objCT.objUserControl.Controls("UpDownRADe")
End Property

Public Property Get UpDownRAAte(ByVal objCT As Object) As Object
     Set UpDownRAAte = objCT.objUserControl.Controls("UpDownRAAte")
End Property

Public Property Get ContaCorrenteSeleciona(ByVal objCT As Object) As Object
     Set ContaCorrenteSeleciona = objCT.objUserControl.Controls("ContaCorrenteSeleciona")
End Property

Public Property Get CtaCorrenteTodas(ByVal objCT As Object) As Object
     Set CtaCorrenteTodas = objCT.objUserControl.Controls("CtaCorrenteTodas")
End Property

Public Property Get CtaCorrenteApenas(ByVal objCT As Object) As Object
     Set CtaCorrenteApenas = objCT.objUserControl.Controls("CtaCorrenteApenas")
End Property

Public Property Get TipoDocSeleciona(ByVal objCT As Object) As Object
     Set TipoDocSeleciona = objCT.objUserControl.Controls("TipoDocSeleciona")
End Property

Public Property Get TipoDocTodos(ByVal objCT As Object) As Object
     Set TipoDocTodos = objCT.objUserControl.Controls("TipoDocTodos")
End Property

Public Property Get TipoDocApenas(ByVal objCT As Object) As Object
     Set TipoDocApenas = objCT.objUserControl.Controls("TipoDocApenas")
End Property

Public Function BR_Inicializa_Grid_Parcelas(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Parcelas

    'tela em questão
    Set objGridInt.objForm = objCT

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Baixar")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Saldo")
    objGridInt.colColuna.Add ("Valor Baixar")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Multa")
    objGridInt.colColuna.Add ("Juros")
    objGridInt.colColuna.Add ("Valor Recebido")
    
    '########################################
    'Inserido por Wagner
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Voucher")
    objGridInt.colColuna.Add ("C.Cred.")
    '########################################
    
    'Visibilidade para versão LIGHT
    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colColuna.Add ("Cobrador")
    End If
    objGridInt.colColuna.Add ("Valor Parcela")
    
    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colColuna.Add ("Filial Empresa")
    End If
    
    'campos de edição do grid
    objGridInt.colCampo.Add (objCT.Selecionar.Name)
    objGridInt.colCampo.Add (objCT.DataVencimento.Name)
    objGridInt.colCampo.Add (objCT.Tipo.Name)
    objGridInt.colCampo.Add (objCT.Numero.Name)
    objGridInt.colCampo.Add (objCT.Parcela.Name)
    objGridInt.colCampo.Add (objCT.Saldo.Name)
    objGridInt.colCampo.Add (objCT.ValorBaixar.Name)
    objGridInt.colCampo.Add (objCT.ValorDesconto.Name)
    objGridInt.colCampo.Add (objCT.ValorMulta.Name)
    objGridInt.colCampo.Add (objCT.ValorJuros.Name)
    objGridInt.colCampo.Add (objCT.ValorAReceber.Name)
    
    '######################################
    'Inserido por Wagner
    objGridInt.colCampo.Add (objCT.ClienteItem.Name)
    objGridInt.colCampo.Add (objCT.FilialClienteItem.Name)
    objGridInt.colCampo.Add (NumVou(objCT).Name)
    objGridInt.colCampo.Add (NumCCred(objCT).Name)
    '######################################
    
    'Visibilidade para versão LIGHT
    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colCampo.Add (objCT.Cobrador.Name)
    End If
    objGridInt.colCampo.Add (objCT.ValorParcela.Name)
    
    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colCampo.Add (objCT.FilialEmpresa.Name)
    End If
    
    objCT.iGrid_Baixar_Col = 1
    objCT.iGrid_DataVencimento_Col = 2
    objCT.iGrid_Tipo_Col = 3
    objCT.iGrid_Numero_Col = 4
    objCT.iGrid_Parcela_Col = 5
    objCT.iGrid_Saldo_Col = 6
    objCT.iGrid_ValorBaixar_Col = 7
    objCT.iGrid_ValorDesconto_Col = 8
    objCT.iGrid_ValorMulta_Col = 9
    objCT.iGrid_ValorJuros_Col = 10
    objCT.iGrid_ValorAReceber_Col = 11
    objCT.iGrid_Cliente_Col = 12
    objCT.iGrid_Filial_Col = 13
    objCT.gobjinfousu.iGrid_NumVou_Col = 14
    objCT.gobjinfousu.iGrid_NumCCred_Col = 15
    
    'Visibilidade para versão LIGHT
    If giTipoVersao = VERSAO_FULL Then
        objCT.iGrid_Cobrador_Col = 16
        objCT.iGrid_ValorParcela_Col = 17
        objCT.iGrid_FilialEmpresa_Col = 18
    ElseIf giTipoVersao = VERSAO_LIGHT Then
        objCT.iGrid_ValorParcela_Col = 16
    End If
    
    objGridInt.objGrid = objCT.GridParcelas

    'todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_PARCELAS_BAIXA + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7

    'largura da primeira coluna
    objCT.GridParcelas.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'incluir barra de rolagem horizontal
    objGridInt.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    'Não permite incluir novas linhas nem excluir as existentes
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    BR_Inicializa_Grid_Parcelas = SUCESSO
    
    Exit Function

End Function

Function BR_Traz_Parcela_Tela(ByVal objCT As Object, ByVal lNumIntParc As Long, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim sNumCCred As String
Dim lNumVou As Long

On Error GoTo Erro_BR_Traz_Parcela_Tela

    lErro = CF("TRPNumCCred_Le_NumIntParc", lNumIntParc, sNumCCred, lNumVou)
    If lErro <> SUCESSO Then gError 182831

    If lNumVou <> 0 Then
        objCT.GridParcelas.TextMatrix(iLinha, objCT.gobjinfousu.iGrid_NumVou_Col) = lNumVou
    Else
        objCT.GridParcelas.TextMatrix(iLinha, objCT.gobjinfousu.iGrid_NumVou_Col) = ""
    End If
    
    objCT.GridParcelas.TextMatrix(iLinha, objCT.gobjinfousu.iGrid_NumCCred_Col) = sNumCCred

    BR_Traz_Parcela_Tela = SUCESSO
     
    Exit Function
    
Erro_BR_Traz_Parcela_Tela:

    BR_Traz_Parcela_Tela = gErr
     
    Select Case gErr
    
        Case 182831
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182832)
     
    End Select
     
    Exit Function
    
End Function

Public Function BR_Inicializa_Grid_RecebAntecipados(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Recebimentos Antecipados

    'tela em questão
    Set objGridInt.objForm = objCT

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Data Movimento")
    objGridInt.colColuna.Add ("Conta Corrente")
    objGridInt.colColuna.Add ("Meio Pagamento")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Saldo")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Doc")
    objGridInt.colColuna.Add ("Seq")
    objGridInt.colColuna.Add ("Histórico")
    objGridInt.colColuna.Add ("Filial Empresa")

   'campos de edição do grid
    objGridInt.colCampo.Add (objCT.SelecionarRA.Name)
    objGridInt.colCampo.Add (objCT.DataMovimentoRA.Name)
    objGridInt.colCampo.Add (objCT.ContaCorrenteRA.Name)
    objGridInt.colCampo.Add (objCT.MeioPagtoRA.Name)
    objGridInt.colCampo.Add (objCT.ValorRA.Name)
    objGridInt.colCampo.Add (objCT.SaldoRA.Name)
    objGridInt.colCampo.Add (ClienteRA(objCT).Name)
    objGridInt.colCampo.Add (FilialClienteRA(objCT).Name)
    objGridInt.colCampo.Add (objCT.Doc.Name)
    objGridInt.colCampo.Add (objCT.Seq.Name)
    objGridInt.colCampo.Add (objCT.Hist.Name)
    objGridInt.colCampo.Add (objCT.FilialRA.Name)

    objCT.iGrid_SelecionarRA_Col = 1
    objCT.iGrid_DataMovimentoRA_Col = 2
    objCT.iGrid_ContaCorrenteRA_Col = 3
    objCT.iGrid_MeioPagtoRA_Col = 4
    objCT.iGrid_ValorRA_Col = 5
    objCT.iGrid_SaldoRA_Col = 6
    objCT.gobjinfousu.iGrid_ClienteRA_Col = 7
    objCT.gobjinfousu.iGrid_FilialClienteRA_Col = 8
    objCT.iGrid_Doc_Col = 9
    objCT.iGrid_Seq_Col = 10
    objCT.iGrid_Hist_Col = 11
    objCT.iGrid_FilialRA_Col = 12

    objGridInt.objGrid = objCT.GridRecebAntecipados

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 3

    'largura da primeira coluna
    objCT.GridRecebAntecipados.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não permite incluir novas linhas nem excluir as existentes
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    BR_Inicializa_Grid_RecebAntecipados = SUCESSO
    
    Exit Function

End Function

Function BR_Traz_RA_Tela(ByVal objCT As Object, ByVal objRecebAntecipados As ClassAntecipReceb, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_BR_Traz_RA_Tela

    objCliente.lCodigo = objRecebAntecipados.lCliente
    objFilialCliente.lCodCliente = objRecebAntecipados.iFilial
    objFilialCliente.iCodFilial = objRecebAntecipados.iFilial
    
    'le o nome reduzido do cliente
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 182871
        
    'le o nome reduzido da filial  cliente
    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 12567 Then gError 182872

    objCT.GridRecebAntecipados.TextMatrix(iLinha, objCT.gobjinfousu.iGrid_ClienteRA_Col) = objRecebAntecipados.lCliente & SEPARADOR & objCliente.sNomeReduzido
    objCT.GridRecebAntecipados.TextMatrix(iLinha, objCT.gobjinfousu.iGrid_FilialClienteRA_Col) = objRecebAntecipados.iFilial & SEPARADOR & objFilialCliente.sNome

    BR_Traz_RA_Tela = SUCESSO
     
    Exit Function
    
Erro_BR_Traz_RA_Tela:

    BR_Traz_RA_Tela = gErr
     
    Select Case gErr
    
        Case 182871, 182872
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182873)
     
    End Select
     
    Exit Function
    
End Function

Public Sub TipoDocApenas_Click(ByVal objCT As Object)

    'Habilita a combo para a seleção da conta corrente
    TipoDocSeleciona(objCT).Enabled = True
    
    objCT.iFramePrincipalAlterado = REGISTRO_ALTERADO
    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDocTodos_Click(ByVal objCT As Object)

    'Desabilita a combo para a seleção da conta corrente
    TipoDocSeleciona(objCT).Enabled = False

    'Limpa a combo de seleção de conta corrente
    TipoDocSeleciona(objCT).ListIndex = COMBO_INDICE

    objCT.iFramePrincipalAlterado = REGISTRO_ALTERADO
    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDocSeleciona_Change(ByVal objCT As Object)

    objCT.iFramePrincipalAlterado = REGISTRO_ALTERADO
    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CtaCorrenteApenas_Click(ByVal objCT As Object)

    'Habilita a combo para a seleção da conta corrente
    ContaCorrenteSeleciona(objCT).Enabled = True
    
    objCT.iFramePrincipalAlterado = REGISTRO_ALTERADO
    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CtaCorrenteTodas_Click(ByVal objCT As Object)

    'Desabilita a combo para a seleção da conta corrente
    ContaCorrenteSeleciona(objCT).Enabled = False

    'Limpa a combo de seleção de conta corrente
    ContaCorrenteSeleciona(objCT).ListIndex = COMBO_INDICE

    objCT.iFramePrincipalAlterado = REGISTRO_ALTERADO
    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaCorrenteSeleciona_Change(ByVal objCT As Object)

    objCT.iFramePrincipalAlterado = REGISTRO_ALTERADO
    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaCorrenteSeleciona_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim sContaCorrenteSeleciona As String
Dim iCodigo As Integer
Dim objContaCorrente As New ClassContasCorrentesInternas

On Error GoTo Erro_ContaCorrenteSeleciona_Validate

    'Se a ComboBox ContaCorrenteSeleciona está preenchida
    If Len(Trim(ContaCorrenteSeleciona(objCT).Text)) <> 0 Then

        'Verifica se não está preenchida com o ítem selecionado na ComboBox ContaCorrenteSeleciona
        If ContaCorrenteSeleciona(objCT).Text <> ContaCorrenteSeleciona(objCT).List(ContaCorrenteSeleciona(objCT).ListIndex) Then
        
            'Verifica se existe o ítem na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(ContaCorrenteSeleciona(objCT), iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 182875
        
            'Nao existe o ítem com o CÓDIGO na List da ComboBox
            If lErro = 6730 Then
        
                'Lê os dados da conta corrente a partir do código passado como parâmetro
                lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrente)
                If lErro <> SUCESSO And lErro <> 11807 Then gError 182876
        
                'Se não encontrou a conta => o número de erro recebe tratamento para perguntar se deseja criar a conta
                If lErro = 11807 Then gError 182877
        
                'Coloca os dados da conta na combo
                ContaCorrenteSeleciona(objCT).Text = CStr(objContaCorrente.iCodigo) & SEPARADOR & objContaCorrente.sNomeReduzido
        
            End If
        
            'Não existe o ítem com a STRING na List da ComboBox
            If lErro = 6731 Then
        
                'Lê os dados da conta corrente a partir do nome red. passado como parâmetro
                objContaCorrente.sNomeReduzido = ContaCorrenteSeleciona(objCT).Text
                lErro = CF("ContaCorrenteInt_Le_NomeReduzido", objContaCorrente)
                If lErro <> SUCESSO And lErro <> 15981 Then gError 182878
        
                'Se não encontrou a conta => Erro
                If lErro = 15981 Then gError 182879
                
                'Coloca os dados da conta na combo
                ContaCorrenteSeleciona(objCT).Text = CStr(objContaCorrente.iCodigo) & SEPARADOR & objContaCorrente.sNomeReduzido
            
            End If
            
        End If
        
    End If

    Exit Sub

Erro_ContaCorrenteSeleciona_Validate:

    Cancel = True

    Select Case gErr

        Case 182875, 182876, 182878

        Case 182877
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONTACORRENTE", ContaCorrenteSeleciona(objCT).Text)

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("CtaCorrenteInt", objContaCorrente)
            End If

        Case 182879
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr, objContaCorrente.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182880)

    End Select

    Exit Sub

End Sub

Public Sub DataRADe_Validate(ByVal objCT As Object, Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataRADe_Validate

    'Se a data DataRADe está preenchida
    If Len(DataRADe(objCT).ClipText) > 0 Then

        'Verifica se a data DataRADe é válida
        lErro = Data_Critica(DataRADe(objCT).Text)
        If lErro <> SUCESSO Then gError 182881

    End If

    Exit Sub

Erro_DataRADe_Validate:

    Cancel = True

    Select Case gErr

        Case 182881

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182882)

    End Select

    Exit Sub

End Sub

Public Sub DataRADe_Change(ByVal objCT As Object)

    'Registra que houve alteração
    objCT.iAlterado = REGISTRO_ALTERADO
    objCT.iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataRADe_GotFocus(ByVal objCT As Object)

Dim iFrameAux As Integer
    
    iFrameAux = objCT.iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataRADe(objCT), objCT.iAlterado)
    objCT.iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub DataRAAte_Validate(ByVal objCT As Object, Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataRAAte_Validate

    'Se a data DataRAAte está preenchida
    If Len(DataRAAte(objCT).ClipText) > 0 Then

        'Verifica se a data DataRAAte é válida
        lErro = Data_Critica(DataRAAte(objCT).Text)
        If lErro <> SUCESSO Then gError 182883

    End If

    Exit Sub

Erro_DataRAAte_Validate:

    Cancel = True

    Select Case gErr

        Case 182883

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182884)

    End Select

    Exit Sub

End Sub

Public Sub DataRAAte_Change(ByVal objCT As Object)

    'Registra que houve alteração
    objCT.iAlterado = REGISTRO_ALTERADO
    objCT.iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataRAAte_GotFocus(ByVal objCT As Object)

Dim iFrameAux As Integer
    
    iFrameAux = objCT.iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataRAAte(objCT), objCT.iAlterado)
    objCT.iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub UpDownRADe_DownClick(ByVal objCT As Object)

Dim lErro As Long

On Error GoTo Erro_UpDownRADe_DownClick

    'Diminui a DataRADe em 1 dia
    lErro = Data_Up_Down_Click(DataRADe(objCT), DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 182885

    Exit Sub

Erro_UpDownRADe_DownClick:

    Select Case gErr

        Case 182885

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182886)

    End Select

    Exit Sub

End Sub

Public Sub UpDownRADe_UpClick(ByVal objCT As Object)

Dim lErro As Long

On Error GoTo Erro_UpDownRADe_UpClick

    'Aumenta a DataRADe em 1 dia
    lErro = Data_Up_Down_Click(DataRADe(objCT), AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 182887

    Exit Sub

Erro_UpDownRADe_UpClick:

    Select Case gErr

        Case 182887

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182888)

    End Select

    Exit Sub

End Sub

Public Sub UpDownRAAte_DownClick(ByVal objCT As Object)

Dim lErro As Long

On Error GoTo Erro_UpDownRAAte_DownClick

    'Diminui a DataRAAte em 1 dia
    lErro = Data_Up_Down_Click(DataRAAte(objCT), DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 182889

    Exit Sub

Erro_UpDownRAAte_DownClick:

    Select Case gErr

        Case 182889

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182890)

    End Select

    Exit Sub

End Sub

Public Sub UpDownRAAte_UpClick(ByVal objCT As Object)

Dim lErro As Long

On Error GoTo Erro_UpDownRAAte_UpClick

    'Aumenta a DataRAAte em 1 dia
    lErro = Data_Up_Down_Click(DataRAAte(objCT), AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 182891

    Exit Sub

Erro_UpDownRAAte_UpClick:

    Select Case gErr

        Case 182891

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182892)

    End Select

    Exit Sub

End Sub

Public Sub ValorDe_Change(ByVal objCT As Object)

    'Registra que houve alteração
    objCT.iAlterado = REGISTRO_ALTERADO
    objCT.iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorAte_Change(ByVal objCT As Object)

    'Registra que houve alteração
    objCT.iAlterado = REGISTRO_ALTERADO
    objCT.iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorDe_Validate(ByVal objCT As Object, Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_ValorDe_Validate

    'Se a data ValorDe está preenchida
    If Len(ValorDe(objCT).Text) > 0 Then

        'Verifica se a data ValorDe é válida
        lErro = Valor_Positivo_Critica_Double(ValorDe(objCT).Text)
        If lErro <> SUCESSO Then gError 182893

    End If

    Exit Sub

Erro_ValorDe_Validate:

    Cancel = True

    Select Case gErr

        Case 182893

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182894)

    End Select

    Exit Sub

End Sub

Public Sub ValorAte_Validate(ByVal objCT As Object, Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_ValorAte_Validate

    'Se a data ValorAte está preenchida
    If Len(ValorAte(objCT).Text) > 0 Then

        'Verifica se a data ValorAte é válida
        lErro = Valor_Positivo_Critica_Double(ValorAte(objCT).Text)
        If lErro <> SUCESSO Then gError 182895

    End If

    Exit Sub

Erro_ValorAte_Validate:

    Cancel = True

    Select Case gErr

        Case 182895

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182896)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Combos_ContaCorrente(objComboBox As ComboBox) As Long

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNomeRed As AdmCodigoNome

On Error GoTo Erro_Carrega_Combos_ContaCorrente

    'Lê Codigos, NomesReduzidos de ContasCorrentes
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then gError 182897

    'Preeche list de ComboBox
    For Each objCodigoNomeRed In colCodigoNomeRed
        objComboBox.AddItem CStr(objCodigoNomeRed.iCodigo) & SEPARADOR & objCodigoNomeRed.sNome
        objComboBox.ItemData(objComboBox.NewIndex) = objCodigoNomeRed.iCodigo
    Next

    Carrega_Combos_ContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_Combos_ContaCorrente:

    Carrega_Combos_ContaCorrente = gErr

    Select Case gErr

        Case 182897

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182898)

    End Select

    Exit Function

End Function

Function BR_Carrega_Tab_Parcelas(ByVal objCT As Object, lCodCli As Long, iFilialCli As Integer, dtEmissaoInic As Date, dtEmissaoFim As Date, dtVencInic As Date, dtVencFim As Date, lTituloInic As Long, lTituloFim As Long, ByVal objFilialCliente As ClassFilialCliente, ByVal gcolInfoParcRec As Collection, ByVal gcolRecebAntecipado As Collection, Optional ByVal lCodViagem As Long, Optional ByVal iFormaPagamento As Integer, Optional ByVal iVendedor As Integer) As Long

Dim lErro As Long
Dim iConta As Integer
Dim sTipoDoc As String

On Error GoTo Erro_BR_Carrega_Tab_Parcelas

    If StrParaDbl(ValorDe(objCT).Text) <> 0 And StrParaDbl(ValorAte(objCT).Text) <> 0 Then
        If StrParaDbl(ValorDe(objCT).Text) > StrParaDbl(ValorAte(objCT).Text) Then gError 182900
    End If
    
    If StrParaDate(DataRADe(objCT).Text) <> DATA_NULA And StrParaDate(DataRAAte(objCT).Text) <> DATA_NULA Then
        If StrParaDate(DataRADe(objCT).Text) > StrParaDate(DataRAAte(objCT).Text) Then gError 182901
    End If
    
    If CtaCorrenteApenas(objCT).Value = True Then
        If ContaCorrenteSeleciona(objCT).ListIndex = -1 Then gError 182902
        iConta = Codigo_Extrai(ContaCorrenteSeleciona(objCT).Text)
    Else
        iConta = 0
    End If
    
    If TipoDocApenas(objCT).Value = True Then
        If TipoDocSeleciona(objCT).ListIndex = -1 Then gError 185038
        sTipoDoc = SCodigo_Extrai(TipoDocSeleciona(objCT).Text)
    Else
        sTipoDoc = ""
    End If

    'Preenche a Coleção de Parcelas
    lErro = CF("ParcelasRec_Le_BaixaManualSel_TRP", lCodCli, iFilialCli, dtEmissaoInic, dtEmissaoFim, dtVencInic, dtVencFim, lTituloInic, lTituloFim, StrParaDbl(ValorDe(objCT).Text), StrParaDbl(ValorAte(objCT).Text), sTipoDoc, gcolInfoParcRec)
    If lErro <> SUCESSO Then gError 182897

    'Preenche a Coleção de Recebimentos antecipados
    lErro = CF("RecebAntecipados_Le_FilialCliente_Baixa_TRP", StrParaDate(DataRADe(objCT).Text), StrParaDate(DataRAAte(objCT).Text), iConta, gcolRecebAntecipado)
    If lErro <> SUCESSO And lErro <> 29142 Then gError 182898
    
    BR_Carrega_Tab_Parcelas = SUCESSO

    Exit Function

Erro_BR_Carrega_Tab_Parcelas:

    BR_Carrega_Tab_Parcelas = gErr

    Select Case gErr
    
        Case 182897, 182898

        Case 182900
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", gErr)

        Case 182901
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR_DATA_FINAL", gErr)
               
        Case 182902
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", gErr)

        Case 185038
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182899)

    End Select

    Exit Function

End Function

Function BR_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_BR_Form_Load

    lErro = Carrega_Combos_ContaCorrente(ContaCorrenteSeleciona(objCT))
    If lErro <> SUCESSO Then gError 185034
    
    lErro = Carrega_TipoDocumento(TipoDocSeleciona(objCT))
    If lErro <> SUCESSO Then gError 185035

    BR_Form_Load = SUCESSO

    Exit Function

Erro_BR_Form_Load:

    BR_Form_Load = gErr

    Select Case gErr
    
        Case 185034, 185035

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185036)

    End Select

    Exit Function
    
End Function

Function BR_Limpa_Tela_BaixaRec(ByVal objCT As Object) As Long
    
    CtaCorrenteTodas(objCT).Value = True
    Call CtaCorrenteTodas_Click(objCT)
    
    TipoDocTodos(objCT).Value = True
    Call TipoDocTodos_Click(objCT)
    
    BR_Limpa_Tela_BaixaRec = SUCESSO
End Function

Private Function Carrega_TipoDocumento(ByVal objComboBox As ComboBox)
'Carrega os Tipos de Documento

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    'Le os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    If lErro <> SUCESSO Then gError 185032
    
    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
                    
        objComboBox.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
    
    Next

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = gErr

    Select Case gErr

        Case 185032

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185033)

    End Select

    Exit Function

End Function



