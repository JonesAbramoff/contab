VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTProdutoTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gobjCT As Object

Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Public Property Get PercComissao(ByVal objCT As Object) As Object
     Set PercComissao = objCT.objUserControl.Controls("PercComissao")
End Property

Public Property Get Vendedor(ByVal objCT As Object) As Object
     Set Vendedor = objCT.objUserControl.Controls("Vendedor")
End Property

Public Property Get DiasDe(ByVal objCT As Object) As Object
     Set DiasDe = objCT.objUserControl.Controls("DiasDe")
End Property

Public Property Get DiasAte(ByVal objCT As Object) As Object
     Set DiasAte = objCT.objUserControl.Controls("DiasAte")
End Property

Public Property Get Diario(ByVal objCT As Object) As Object
     Set Diario = objCT.objUserControl.Controls("Diario")
End Property

Public Property Get ValorTarifa(ByVal objCT As Object) As Object
     Set ValorTarifa = objCT.objUserControl.Controls("ValorTarifa")
End Property

Public Property Get ValorTarifaAdicional(ByVal objCT As Object) As Object
     Set ValorTarifaAdicional = objCT.objUserControl.Controls("ValorTarifaAdicional")
End Property

Public Property Get GridTarifa(ByVal objCT As Object) As Object
     Set GridTarifa = objCT.objUserControl.Controls("GridTarifa")
End Property

Public Property Get DataVigencia(ByVal objCT As Object) As Object
     Set DataVigencia = objCT.objUserControl.Controls("DataVigencia")
End Property

Public Property Get IncluirTarifas(ByVal objCT As Object) As Object
     Set IncluirTarifas = objCT.objUserControl.Controls("IncluirTarifas")
End Property

Public Property Get DescontoTitular(ByVal objCT As Object) As Object
     Set DescontoTitular = objCT.objUserControl.Controls("DescontoTitular")
End Property

Public Property Get DescontoAcompanhante(ByVal objCT As Object) As Object
     Set DescontoAcompanhante = objCT.objUserControl.Controls("DescontoAcompanhante")
End Property

Public Property Get MaxAcompanhantes(ByVal objCT As Object) As Object
     Set MaxAcompanhantes = objCT.objUserControl.Controls("MaxAcompanhantes")
End Property

Public Property Get Moeda(ByVal objCT As Object) As Object
     Set Moeda = objCT.objUserControl.Controls("Moeda")
End Property

Public Property Get IdadeLimite(ByVal objCT As Object) As Object
     Set IdadeLimite = objCT.objUserControl.Controls("IdadeLimite")
End Property

Public Property Get ExibeWeb(ByVal objCT As Object) As Object
     Set ExibeWeb = objCT.objUserControl.Controls("ExibeWeb")
End Property

Public Property Get AceitaIdadeMaiorLimite(ByVal objCT As Object) As Object
     Set AceitaIdadeMaiorLimite = objCT.objUserControl.Controls("AceitaIdadeMaiorLimite")
End Property

Public Property Get PercAcrescIdadeMaiorLimite(ByVal objCT As Object) As Object
     Set PercAcrescIdadeMaiorLimite = objCT.objUserControl.Controls("PercAcrescIdadeMaiorLimite")
End Property

Public Property Get VigenciaFixa(ByVal objCT As Object) As Object
     Set VigenciaFixa = objCT.objUserControl.Controls("VigenciaFixa")
End Property

Public Property Get AceitaTodosDestinos(ByVal objCT As Object) As Object
     Set AceitaTodosDestinos = objCT.objUserControl.Controls("AceitaTodosDestinos")
End Property

Public Property Get GridDestinos(ByVal objCT As Object) As Object
     Set GridDestinos = objCT.objUserControl.Controls("GridDestinos")
End Property

Public Property Get Destino(ByVal objCT As Object) As Object
     Set Destino = objCT.objUserControl.Controls("Destino")
End Property

Public Property Get TextoVouPT(ByVal objCT As Object) As Object
     Set TextoVouPT = objCT.objUserControl.Controls("TextoVouPT")
End Property

Public Property Get TextoVouEN(ByVal objCT As Object) As Object
     Set TextoVouEN = objCT.objUserControl.Controls("TextoVouEN")
End Property

Public Property Get FrameDestino(ByVal objCT As Object) As Object
     Set FrameDestino = objCT.objUserControl.Controls("FrameDestino")
End Property

Public Function Prod_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Prod_Form_Load

    Set objCT.gobjInfoUsu.objGridTarifa = New AdmGrid
    Set objCT.gobjInfoUsu.objGridDestinos = New AdmGrid
    Set objCT.gobjInfoUsu.gobjProduto = New ClassProduto
    
    'carrega a combo de Moedas
    lErro = Carrega_Moeda(objCT)
    If lErro <> SUCESSO Then gError 200285

    lErro = Inicializa_Grid_Tarifa(objCT, objCT.gobjInfoUsu.objGridTarifa)
    If lErro <> SUCESSO Then gError 200286
    
    lErro = Inicializa_Grid_Destinos(objCT, objCT.gobjInfoUsu.objGridDestinos)
    If lErro <> SUCESSO Then gError 200286
    
    'Carrega a combo Destino
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_DESTINO_VIAGEM, Destino(objCT))
    If lErro <> SUCESSO Then gError 200286
        
    ExibeWeb(objCT).Value = vbChecked
    IncluirTarifas(objCT).Value = vbUnchecked

    Prod_Form_Load = SUCESSO
         
    Exit Function
    
Erro_Prod_Form_Load:

    Prod_Form_Load = gErr
     
    Select Case gErr
    
        Case 200285, 200286

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200287)
     
    End Select
     
    Exit Function
    
End Function

Public Function Prod_Form_UnLoad(ByVal objCT As Object) As Long

    Set objEventoVendedor = Nothing
    Set gobjCT = Nothing
    
    Set objCT.gobjInfoUsu.objGridTarifa = Nothing
    Set objCT.gobjInfoUsu.gobjProduto = Nothing

    Prod_Form_UnLoad = SUCESSO
    
End Function

Public Function Prod_Gravar_Registro(ByVal objCT As Object, objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProdutoTRP As New ClassProdutoTRP
Dim objVendedor As New ClassVendedor
Dim objTarifa As ClassTRPProdTarifa
Dim iDiaAteAnt As Integer

On Error GoTo Erro_Prod_Gravar_Registro

    If Len(Trim(Vendedor(objCT).Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor(objCT).Text

    'Verifica se vendedor existe
    If objVendedor.sNomeReduzido <> "" Then
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 197229

        'Não encontrou o vendedor ==> erro
        If lErro = 25008 Then gError 197230

        objProdutoTRP.iVendedor = objVendedor.iCodigo

    End If
    
    If Len(Trim(PercComissao(objCT).Text)) > 0 Then objProdutoTRP.dPercComissao = CDbl(PercComissao(objCT).Text) / 100

    objProdutoTRP.dDescontoAcompanhante = StrParaDbl(DescontoAcompanhante(objCT).Text) / 100
    objProdutoTRP.dDescontoTitular = StrParaDbl(DescontoTitular(objCT).Text) / 100
    objProdutoTRP.dPercAcrescIdadeMaiorLimite = StrParaDbl(PercAcrescIdadeMaiorLimite(objCT).Text) / 100
    
    If ExibeWeb(objCT).Value = vbChecked Then
        objProdutoTRP.iExibeWeb = MARCADO
    Else
        objProdutoTRP.iExibeWeb = DESMARCADO
    End If
    
    If AceitaIdadeMaiorLimite(objCT).Value = vbChecked Then
        objProdutoTRP.iAceitaIdadeMaiorLimite = MARCADO
    Else
        objProdutoTRP.iAceitaIdadeMaiorLimite = DESMARCADO
    End If
    
    objProdutoTRP.sProduto = objProduto.sCodigo
    objProdutoTRP.iLimiteIdade = StrParaInt(IdadeLimite(objCT).Text)
    objProdutoTRP.iMaxAcompanhante = StrParaInt(MaxAcompanhantes(objCT).Text)
    objProdutoTRP.iMoeda = Codigo_Extrai(Moeda(objCT).Text)
    objProdutoTRP.iVigenciaFixa = StrParaInt(VigenciaFixa(objCT).Text)
    
    If AceitaTodosDestinos(objCT).Value = vbChecked Then
        objProdutoTRP.iTodosDestValidos = MARCADO
    Else
        objProdutoTRP.iTodosDestValidos = DESMARCADO
        For iIndice = 1 To objCT.gobjInfoUsu.objGridDestinos.iLinhasExistentes
            If Codigo_Extrai(GridDestinos(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Destino_Col)) <> 0 Then
                objProdutoTRP.colDestinos.Add Codigo_Extrai(GridDestinos(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Destino_Col))
            End If
        Next
    End If
    objProdutoTRP.sTextoVouEN = TextoVouEN(objCT).Text
    objProdutoTRP.sTextoVouPT = TextoVouPT(objCT).Text

    If IncluirTarifas(objCT).Value = vbChecked Then

        objProdutoTRP.dtDataVigencia = StrParaDate(DataVigencia(objCT).Text)
        
        If objProdutoTRP.dtDataVigencia < Date Then gError 200288
        
        'Verifica se no grid de tarifa as informacoes estao preenchidas
        For iIndice = 1 To objCT.gobjInfoUsu.objGridTarifa.iLinhasExistentes
        
            Set objTarifa = New ClassTRPProdTarifa

            objTarifa.dtVigencia = objProdutoTRP.dtDataVigencia
            objTarifa.dValor = StrParaDbl(GridTarifa(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ValorTarifa_Col))
            objTarifa.dValorAdicional = StrParaDbl(GridTarifa(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ValorTarifaAdicional_Col))
            objTarifa.iDiasAte = StrParaInt(GridTarifa(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_DiasAte_Col))
            objTarifa.iDiasDe = StrParaInt(GridTarifa(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_DiasDe_Col))
            objTarifa.sProduto = objProduto.sCodigo
            
            If StrParaInt(GridTarifa(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Diario_Col)) = MARCADO Then
                objTarifa.iDiario = MARCADO
            Else
                objTarifa.iDiario = DESMARCADO
            End If
            
            If objTarifa.iDiasAte = 0 Then gError 200289
            If objTarifa.iDiasDe > objTarifa.iDiasAte Then gError 200290
            If objTarifa.dValor = 0 Then gError 200291
            If objTarifa.iDiario = MARCADO And objTarifa.dValorAdicional = 0 Then gError 200292
            If objTarifa.iDiasDe <= iDiaAteAnt Then gError 200295
            
            iDiaAteAnt = objTarifa.iDiasAte
            
            objProdutoTRP.colTarifas.Add objTarifa
             
        Next
    
    End If

    Set objProduto.objInfoUsu = objProdutoTRP

    Prod_Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Prod_Gravar_Registro:
    
    Prod_Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 197229
        
        Case 197230
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objVendedor.sNomeReduzido)
        
        Case 200288
            Call Rotina_Erro(vbOKOnly, "ERRO_VIGENCIA_INFERIOR_DATA_ATUAL", gErr)
        
        Case 200289
            Call Rotina_Erro(vbOKOnly, "ERRO_VIG_ATE_NAO_PREENCH_GRID", gErr, iIndice)
        
        Case 200290
            Call Rotina_Erro(vbOKOnly, "ERRO_VIG_DE_MAIOR_VIG_ATE", gErr, iIndice)
        
        Case 200291
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO_GRID", gErr, iIndice)
        
        Case 200292
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORADI_NAO_PREENCHIDO_GRID", gErr, iIndice)
            
        Case 200295
            Call Rotina_Erro(vbOKOnly, "ERRO_VIG_DE_MENOR_VIG_ATE_ANT", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197231)

    End Select

    Exit Function

End Function

Public Function Prod_Traz_Produto_Tela(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProdutoTRP As ClassProdutoTRP
Dim objVendedor As New ClassVendedor
Dim dtDataAnt As Date
Dim iDestino As Variant

On Error GoTo Erro_Prod_Traz_Produto_Tela

    lErro = CF("Produto_Le_Customizado", objProduto)
    If lErro <> SUCESSO Then gError 197232
    
    If Not (objProduto.objInfoUsu Is Nothing) Then
    
        Set objProdutoTRP = objProduto.objInfoUsu
        
        If objProdutoTRP.iVendedor <> 0 Then
        
            Vendedor(objCT).Text = objProdutoTRP.iVendedor
        
            lErro = TP_Vendedor_Le(Vendedor(objCT), objVendedor)
            If lErro <> SUCESSO Then gError 197233
        
        End If
        
        PercComissao(objCT).Text = CStr(100 * objProdutoTRP.dPercComissao)
        
        DescontoAcompanhante(objCT).Text = CStr(100 * objProdutoTRP.dDescontoAcompanhante)
        DescontoTitular(objCT).Text = CStr(100 * objProdutoTRP.dDescontoTitular)
        PercAcrescIdadeMaiorLimite(objCT).Text = CStr(100 * objProdutoTRP.dPercAcrescIdadeMaiorLimite)
            
        If objProdutoTRP.iExibeWeb = MARCADO Then
            ExibeWeb(objCT).Value = vbChecked
        Else
            ExibeWeb(objCT).Value = vbUnchecked
        End If
        
        If objProdutoTRP.iAceitaIdadeMaiorLimite = MARCADO Then
            AceitaIdadeMaiorLimite(objCT).Value = vbChecked
        Else
            AceitaIdadeMaiorLimite(objCT).Value = vbUnchecked
        End If
        
        IdadeLimite(objCT).PromptInclude = False
        IdadeLimite(objCT).Text = CStr(objProdutoTRP.iLimiteIdade)
        IdadeLimite(objCT).PromptInclude = True
    
        VigenciaFixa(objCT).PromptInclude = False
        If objProdutoTRP.iVigenciaFixa <> 0 Then
            VigenciaFixa(objCT).Text = CStr(objProdutoTRP.iVigenciaFixa)
        Else
            VigenciaFixa(objCT).Text = ""
        End If
        VigenciaFixa(objCT).PromptInclude = True
        
        TextoVouEN(objCT).Text = objProdutoTRP.sTextoVouEN
        TextoVouPT(objCT).Text = objProdutoTRP.sTextoVouPT
    
        MaxAcompanhantes(objCT).PromptInclude = False
        MaxAcompanhantes(objCT).Text = CStr(objProdutoTRP.iMaxAcompanhante)
        MaxAcompanhantes(objCT).PromptInclude = True
        
        Call Combo_Seleciona_ItemData(Moeda(objCT), objProdutoTRP.iMoeda)
        
        Call Grid_Limpa(objCT.gobjInfoUsu.objGridDestinos)
        If objProdutoTRP.iTodosDestValidos = DESMARCADO Then
            AceitaTodosDestinos(objCT).Value = vbUnchecked
            iIndice = 0
            For Each iDestino In objProdutoTRP.colDestinos
                iIndice = iIndice + 1
                Call Combo_Seleciona_ItemData(Destino(objCT), iDestino)
                
                GridDestinos(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Destino_Col) = Destino(objCT).Text
            
            Next
        Else
            AceitaTodosDestinos(objCT).Value = vbChecked
        End If
        
        'Atualiza o número de linhas existentes
        objCT.gobjInfoUsu.objGridDestinos.iLinhasExistentes = iIndice
    
    End If
    
    Set objCT.gobjInfoUsu.gobjProduto = objProduto
    
    lErro = CF("Carrega_Combo_Data", DataVigencia(objCT), "TRPProdTarifa", "DataVigencia", "Produto =  '" & objProduto.sCodigo & "'")
    If lErro <> SUCESSO Then gError 200289
    
    dtDataAnt = DATA_NULA
    For iIndice = 0 To DataVigencia(objCT).ListCount - 1
    
        If StrParaDate(DataVigencia(objCT).List(iIndice)) > Date Then Exit For
    
        dtDataAnt = StrParaDate(DataVigencia(objCT).List(iIndice))
    
    Next
    
    If dtDataAnt <> DATA_NULA Then
        DataVigencia(objCT).Text = Format(dtDataAnt, "dd/mm/yyyy")
        Call DataVigencia_Click(objCT)
    End If
    
    Prod_Traz_Produto_Tela = SUCESSO
     
    Exit Function
    
Erro_Prod_Traz_Produto_Tela:

    Prod_Traz_Produto_Tela = gErr
     
    Select Case gErr
    
        Case 197232, 197233, 200289
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197234)
     
    End Select
     
    Exit Function

End Function

Public Sub UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If objCT.objUserControl.ActiveControl Is Vendedor(objCT) Then
            Call LabelVendedor_Click(objCT)
        End If
    
    End If

End Sub

Public Sub Vendedor_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Vendedor_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    If Len(Trim(Vendedor(objCT).Text)) > 0 Then

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le(Vendedor(objCT), objVendedor)
        If lErro <> SUCESSO Then gError 197226
        
        If objVendedor.iAtivo = DESMARCADO Then gError 197227

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 197226
        
        Case 197227
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INATIVO", gErr, objVendedor.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197228)
    
    End Select

End Sub

Public Sub LabelVendedor_Click(ByVal objCT As Object)

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Vendedor(objCT).Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor(objCT).Text
    
    Set gobjCT = objCT

    Set objEventoVendedor = New AdmEvento
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1

    'Preenche campo Vendedor
    Vendedor(gobjCT).Text = objVendedor.sNomeReduzido

    gobjCT.objUserControl.Show

    Exit Sub

End Sub

Public Sub PercComissao_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercComissao_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComissao_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComissao(objCT).Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComissao(objCT).Text)
    If lErro <> SUCESSO Then Error 195849

    'Formata
    PercComissao(objCT).Text = Format(PercComissao(objCT).Text, "Fixed")

    Exit Sub

Erro_PercComissao_Validate:

    Cancel = True


    Select Case Err

        Case 195849
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 195850)
    
    End Select

    Exit Sub

End Sub


Public Sub DescontoAcompanhante_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DescontoAcompanhante_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DescontoAcompanhante_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(DescontoAcompanhante(objCT).Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(DescontoAcompanhante(objCT).Text)
    If lErro <> SUCESSO Then Error 195849

    'Formata
    DescontoAcompanhante(objCT).Text = Format(DescontoAcompanhante(objCT).Text, "Fixed")

    Exit Sub

Erro_DescontoAcompanhante_Validate:

    Cancel = True


    Select Case Err

        Case 195849
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 195850)
    
    End Select

    Exit Sub

End Sub

Public Sub DescontoTitular_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DescontoTitular_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DescontoTitular_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(DescontoTitular(objCT).Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(DescontoTitular(objCT).Text)
    If lErro <> SUCESSO Then Error 195849

    'Formata
    DescontoTitular(objCT).Text = Format(DescontoTitular(objCT).Text, "Fixed")

    Exit Sub

Erro_DescontoTitular_Validate:

    Cancel = True


    Select Case Err

        Case 195849
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 195850)
    
    End Select

    Exit Sub

End Sub

Function Carrega_Moeda(ByVal objCT As Object)

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError 190666
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 190667
    
    For Each objMoeda In colMoedas
    
        Moeda(objCT).AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
        Moeda(objCT).ItemData(Moeda(objCT).NewIndex) = objMoeda.iCodigo
    
    Next

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 190666
        
        Case 190667
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190668)
    
    End Select

End Function

Public Sub Moeda_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub IdadeLimite_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MaxAcompanhantes_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub GridTarifa_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjInfoUsu.objGridTarifa, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridTarifa, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridTarifa_EnterCell(ByVal objCT As Object)
    Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridTarifa, objCT.iAlterado)
End Sub

Public Sub GridTarifa_GotFocus(ByVal objCT As Object)
    Call Grid_Recebe_Foco(objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub GridTarifa_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub GridTarifa_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjInfoUsu.objGridTarifa, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridTarifa, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridTarifa_LeaveCell(ByVal objCT As Object)
    Call objCT.Saida_Celula(objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub GridTarifa_Validate(ByVal objCT As Object, Cancel As Boolean)
    Call Grid_Libera_Foco(objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub GridTarifa_RowColChange(ByVal objCT As Object)
    Call Grid_RowColChange(objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub GridTarifa_Scroll(ByVal objCT As Object)
    Call Grid_Scroll(objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub DiasDe_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DiasDe_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub DiasDe_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub DiasDe_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DiasDe_Validate

    Set objCT.gobjInfoUsu.objGridTarifa.objControle = DiasDe(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridTarifa)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_DiasDe_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Sub DiasAte_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DiasAte_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub DiasAte_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub DiasAte_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DiasAte_Validate

    Set objCT.gobjInfoUsu.objGridTarifa.objControle = DiasAte(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridTarifa)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_DiasAte_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Sub Diario_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Diario_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub Diario_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub Diario_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Diario_Validate

    Set objCT.gobjInfoUsu.objGridTarifa.objControle = Diario(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridTarifa)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_Diario_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Sub ValorTarifa_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorTarifa_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub ValorTarifa_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub ValorTarifa_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorTarifa_Validate

    Set objCT.gobjInfoUsu.objGridTarifa.objControle = ValorTarifa(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridTarifa)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_ValorTarifa_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Sub ValorTarifaAdicional_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorTarifaAdicional_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub ValorTarifaAdicional_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridTarifa)
End Sub

Public Sub ValorTarifaAdicional_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorTarifaAdicional_Validate

    Set objCT.gobjInfoUsu.objGridTarifa.objControle = ValorTarifaAdicional(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridTarifa)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_ValorTarifaAdicional_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Function Prod_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Prod_Saida_Celula

    'Verifica qual o Grid em questão
    Select Case objGridInt.objGrid.Name

        'Se for o GridTarifa
        Case GridTarifa(objCT).Name

            lErro = Saida_Celula_GridTarifa(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195770

        'Se for o GridDestinos
        Case GridDestinos(objCT).Name

            lErro = Saida_Celula_GridDestinos(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195770

    End Select
    
    Prod_Saida_Celula = SUCESSO

    Exit Function

Erro_Prod_Saida_Celula:

    Prod_Saida_Celula = gErr

    Select Case gErr

        Case 195770, 195771

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195772)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridTarifa(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridTarifa

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for Dias De
        Case objCT.gobjInfoUsu.iGrid_DiasDe_Col
            lErro = Saida_Celula_Inteiro(objCT, objGridInt, DiasDe(objCT))
            If lErro <> SUCESSO Then gError 195760

        'Se for Dias Até
        Case objCT.gobjInfoUsu.iGrid_DiasAte_Col
            lErro = Saida_Celula_Inteiro(objCT, objGridInt, DiasAte(objCT))
            If lErro <> SUCESSO Then gError 195761

        'Se for Valor Tarifa
        Case objCT.gobjInfoUsu.iGrid_ValorTarifa_Col
            lErro = Saida_Celula_Valor(objCT, objGridInt, ValorTarifa(objCT))
            If lErro <> SUCESSO Then gError 195761

        'Se for Valor Tafifa Adicional
        Case objCT.gobjInfoUsu.iGrid_ValorTarifaAdicional_Col
            lErro = Saida_Celula_Valor(objCT, objGridInt, ValorTarifaAdicional(objCT))
            If lErro <> SUCESSO Then gError 195761

    End Select

    Saida_Celula_GridTarifa = SUCESSO

    Exit Function

Erro_Saida_Celula_GridTarifa:

    Saida_Celula_GridTarifa = gErr

    Select Case gErr

        Case 195760, 195761

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195762)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridDestinos(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridDestinos

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for Destino
        Case objCT.gobjInfoUsu.iGrid_Destino_Col
            lErro = Saida_Celula_Padrao(objCT, objGridInt, Destino(objCT))
            If lErro <> SUCESSO Then gError 195760


    End Select

    Saida_Celula_GridDestinos = SUCESSO

    Exit Function

Erro_Saida_Celula_GridDestinos:

    Saida_Celula_GridDestinos = gErr

    Select Case gErr

        Case 195760

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195762)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(ByVal objCT As Object, ByVal objGridInt As AdmGrid, ByVal objControle As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = objControle
    
    'Se Desconto está preenchido
    If Len(Trim(objControle.Text)) <> 0 Then

        'Verifica se o Desconto é válido
        lErro = Valor_NaoNegativo_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 195918
        
        objControle.Text = Format(objControle.Text, "STANDARD")
        
        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195919

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 195918, 195919
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195920)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Inteiro(ByVal objCT As Object, ByVal objGridInt As AdmGrid, ByVal objControle As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Inteiro

    Set objGridInt.objControle = objControle
    
    'Se Desconto está preenchido
    If Len(Trim(objControle.Text)) <> 0 Then

        'Verifica se o Desconto é válido
        lErro = Valor_Inteiro_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 195918
        
        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195919

    Saida_Celula_Inteiro = SUCESSO

    Exit Function

Erro_Saida_Celula_Inteiro:

    Saida_Celula_Inteiro = gErr

    Select Case gErr

        Case 195918, 195919
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195920)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Padrao(ByVal objCT As Object, ByVal objGridInt As AdmGrid, ByVal objControle As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle
    
    'Se Desconto está preenchido
    If Len(Trim(objControle.Text)) <> 0 Then
        
        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195919

    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 195919
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195920)

    End Select

    Exit Function

End Function

Public Function Prod_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iLocalChamada As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Prod_Rotina_Grid_Enable

    Select Case objControl.Name

        Case DiasDe(objCT).Name, DiasAte(objCT).Name, Diario(objCT).Name, ValorTarifa(objCT).Name
            If IncluirTarifas(objCT).Value = vbChecked Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case ValorTarifaAdicional(objCT).Name
            If IncluirTarifas(objCT).Value = vbChecked And StrParaInt(GridTarifa(objCT).TextMatrix(GridTarifa(objCT).Row, objCT.gobjInfoUsu.iGrid_Diario_Col)) = MARCADO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
    End Select
    
    Prod_Rotina_Grid_Enable = SUCESSO

    Exit Function

Erro_Prod_Rotina_Grid_Enable:

    Prod_Rotina_Grid_Enable = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195804)

    End Select

    Exit Function

End Function

Public Sub IncluirTarifas_Click(ByVal objCT As Object)

On Error GoTo Erro_IncluirTarifas_Click

    If StrParaDate(DataVigencia(objCT)) < Date And IncluirTarifas(objCT).Value = vbChecked Then
        IncluirTarifas(objCT).Value = vbUnchecked
        gError 200280
    End If

    Exit Sub

Erro_IncluirTarifas_Click:
    
    Select Case gErr
    
        Case 200280
            Call Rotina_Erro(vbOKOnly, "ERRO_VIGENCIA_INFERIOR_DATA_ATUAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200281)
    
    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Tarifa(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Vigência de")
    objGridInt.colColuna.Add ("Até")
    objGridInt.colColuna.Add ("Aceita Diária")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Valor Diária")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (DiasDe(objCT).Name)
    objGridInt.colCampo.Add (DiasAte(objCT).Name)
    objGridInt.colCampo.Add (Diario(objCT).Name)
    objGridInt.colCampo.Add (ValorTarifa(objCT).Name)
    objGridInt.colCampo.Add (ValorTarifaAdicional(objCT).Name)

    'Colunas do Grid
    objCT.gobjInfoUsu.iGrid_DiasDe_Col = 1
    objCT.gobjInfoUsu.iGrid_DiasAte_Col = 2
    objCT.gobjInfoUsu.iGrid_Diario_Col = 3
    objCT.gobjInfoUsu.iGrid_ValorTarifa_Col = 4
    objCT.gobjInfoUsu.iGrid_ValorTarifaAdicional_Col = 5

    'Grid do GridInterno
    objGridInt.objGrid = GridTarifa(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 9

    'Largura da primeira coluna
    GridTarifa(objCT).ColWidth(0) = 300

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Tarifa = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Destinos(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Destino")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Destino(objCT).Name)

    'Colunas do Grid
    objCT.gobjInfoUsu.iGrid_Destino_Col = 1

    'Grid do GridInterno
    objGridInt.objGrid = GridDestinos(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridDestinos(objCT).ColWidth(0) = 300

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Destinos = SUCESSO

    Exit Function

End Function

Function Prod_Limpa_Tela(ByVal objCT As Object) As Long
    
    Set objCT.gobjInfoUsu.gobjProduto = New ClassProduto

    AceitaIdadeMaiorLimite(objCT).Value = vbUnchecked
    ExibeWeb(objCT).Value = vbChecked
    IncluirTarifas(objCT).Value = vbUnchecked
    
    objCT.gobjInfoUsu.sDataAnt = ""

    Call Grid_Limpa(objCT.gobjInfoUsu.objGridTarifa)
    Call Grid_Limpa(objCT.gobjInfoUsu.objGridDestinos)
    
    FrameDestino(objCT).Enabled = False
    AceitaTodosDestinos(objCT).Value = vbChecked
    
    DataVigencia(objCT).Clear
    
    Moeda(objCT).ListIndex = -1

    Prod_Limpa_Tela = SUCESSO
    
End Function

Public Sub DataVigencia_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataVigencia_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataVigencia_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataVigencia(objCT).Text)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataVigencia(objCT).Text)
    If lErro <> SUCESSO Then gError 58475
    
    Call DataVigencia_Click(objCT)
    
    Exit Sub

Erro_DataVigencia_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 58475

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158306)

    End Select

    Exit Sub

End Sub

Public Sub DataVigencia_Click(ByVal objCT As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim objProdutoTRP As ClassProdutoTRP
Dim objTarifa As ClassTRPProdTarifa
Dim iIndice As Integer
Dim vbMsgResult As VbMsgBoxResult

On Error GoTo Erro_DataVigencia_Click

    If objCT.gobjInfoUsu.sDataAnt <> DataVigencia(objCT).Text Then
    
        If IncluirTarifas(objCT).Value = vbChecked Then
            vbMsgResult = Rotina_Aviso(vbYesNo, "AVISO_TROCA_VIGENCIA_LIMPA_TARIFAS")
            If vbMsgResult = vbNo Then gError 200292
        End If
    
        Call Grid_Limpa(objCT.gobjInfoUsu.objGridTarifa)
    
        'Verifica se foi preenchida a ComboBox Data de Vigencia
        If Len(Trim(DataVigencia(objCT).Text)) > 0 Then
    
            If StrParaDate(DataVigencia(objCT).Text) < Date Then IncluirTarifas(objCT).Value = vbUnchecked
            
            Set objProduto = objCT.gobjInfoUsu.gobjProduto
            Set objProdutoTRP = objProduto.objInfoUsu
            
            If Not (objProdutoTRP Is Nothing) Then
            
                iIndice = 0
                For Each objTarifa In objProdutoTRP.colTarifas
                    
                    If objTarifa.dtVigencia = StrParaDate(DataVigencia(objCT).Text) Then
                        iIndice = iIndice + 1
                        GridTarifa(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_DiasDe_Col) = objTarifa.iDiasDe
                        GridTarifa(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_DiasAte_Col) = objTarifa.iDiasAte
                        GridTarifa(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Diario_Col) = objTarifa.iDiario
                        GridTarifa(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ValorTarifa_Col) = Format(objTarifa.dValor, "STANDARD")
                        
                        If objTarifa.dValorAdicional <> 0 Then
                            GridTarifa(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ValorTarifaAdicional_Col) = Format(objTarifa.dValorAdicional, "STANDARD")
                        Else
                            GridTarifa(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ValorTarifaAdicional_Col) = ""
                        End If
                    End If
                
                Next
                
                'Atualiza o número de linhas existentes
                objCT.gobjInfoUsu.objGridTarifa.iLinhasExistentes = iIndice
                
                Call Grid_Refresh_Checkbox(objCT.gobjInfoUsu.objGridTarifa)
                
            End If
        
        End If
        
        objCT.gobjInfoUsu.sDataAnt = DataVigencia(objCT).Text
        
    End If

    Exit Sub
    
Erro_DataVigencia_Click:

    Select Case gErr
    
        Case 200292
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158290)

    End Select

    Exit Sub
        
End Sub

Public Sub ExibeWeb_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub AceitaIdadeMaiorLimite_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercAcrescIdadeMaiorLimite_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercAcrescIdadeMaiorLimite_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercAcrescIdadeMaiorLimite_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercAcrescIdadeMaiorLimite(objCT).Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercAcrescIdadeMaiorLimite(objCT).Text)
    If lErro <> SUCESSO Then Error 195849

    'Formata
    PercAcrescIdadeMaiorLimite(objCT).Text = Format(PercAcrescIdadeMaiorLimite(objCT).Text, "Fixed")

    Exit Sub

Erro_PercAcrescIdadeMaiorLimite_Validate:

    Cancel = True


    Select Case Err

        Case 195849
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 195850)
    
    End Select

    Exit Sub

End Sub

Public Sub GridDestinos_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjInfoUsu.objGridDestinos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridDestinos, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridDestinos_EnterCell(ByVal objCT As Object)
    Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridDestinos, objCT.iAlterado)
End Sub

Public Sub GridDestinos_GotFocus(ByVal objCT As Object)
    Call Grid_Recebe_Foco(objCT.gobjInfoUsu.objGridDestinos)
End Sub

Public Sub GridDestinos_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjInfoUsu.objGridDestinos)
End Sub

Public Sub GridDestinos_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjInfoUsu.objGridDestinos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridDestinos, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridDestinos_LeaveCell(ByVal objCT As Object)
    Call objCT.Saida_Celula(objCT.gobjInfoUsu.objGridDestinos)
End Sub

Public Sub GridDestinos_Validate(ByVal objCT As Object, Cancel As Boolean)
    Call Grid_Libera_Foco(objCT.gobjInfoUsu.objGridDestinos)
End Sub

Public Sub GridDestinos_RowColChange(ByVal objCT As Object)
    Call Grid_RowColChange(objCT.gobjInfoUsu.objGridDestinos)
End Sub

Public Sub GridDestinos_Scroll(ByVal objCT As Object)
    Call Grid_Scroll(objCT.gobjInfoUsu.objGridDestinos)
End Sub

Public Sub Destino_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Destino_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridDestinos)
End Sub

Public Sub Destino_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridDestinos)
End Sub

Public Sub Destino_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Destino_Validate

    Set objCT.gobjInfoUsu.objGridDestinos.objControle = Destino(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridDestinos)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_Destino_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Sub VigenciaFixa_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TextoVouPT_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TextoVouEN_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Function AceitaTodosDestinos_Click(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_AceitaTodosDestinos_Click
    
    objCT.iAlterado = REGISTRO_ALTERADO

    If AceitaTodosDestinos(objCT).Value = vbChecked Then
        Call Grid_Limpa(objCT.gobjInfoUsu.objGridDestinos)
        FrameDestino(objCT).Enabled = False
    Else
        FrameDestino(objCT).Enabled = True
    End If

    AceitaTodosDestinos_Click = SUCESSO
         
    Exit Function
    
Erro_AceitaTodosDestinos_Click:

    AceitaTodosDestinos_Click = gErr
     
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200287)
     
    End Select
     
    Exit Function
    
End Function

Public Function VigenciaFixa_Validade(ByVal objCT As Object, Cancel As Boolean) As Long

Dim lErro As Long
Dim objGridInt As AdmGrid
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_VigenciaFixa_Validade
    
'    Set objGridInt = objCT.gobjInfoUsu.objGridDestinos
'
'    If StrParaInt(VigenciaFixa.Text) <> 0 Then
'
'        objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
'        objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
'
'        If objGridInt.iLinhasExistentes > 1 Then
'            vbResult = Rotina_Aviso(vbYesNo, "AVISO_GRID_LIMPA")
'        End If
'
'    Else
'
'        objGridInt.iProibidoExcluir = 0
'        objGridInt.iProibidoIncluir = 0
'
'    End If

    VigenciaFixa_Validade = SUCESSO
         
    Exit Function
    
Erro_VigenciaFixa_Validade:

    Cancel = True

    VigenciaFixa_Validade = gErr
     
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200287)
     
    End Select
     
    Exit Function
    
End Function
