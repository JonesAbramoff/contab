VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTVendedoresTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gobjCT As Object

Private WithEvents objEventoSuperior As AdmEvento
Attribute objEventoSuperior.VB_VarHelpID = -1
Private WithEvents objEventoRegiao As AdmEvento
Attribute objEventoRegiao.VB_VarHelpID = -1

Public Property Get Cargo(ByVal objCT As Object) As Object
     Set Cargo = objCT.objUserControl.Controls("Cargo")
End Property

Public Property Get Superior(ByVal objCT As Object) As Object
     Set Superior = objCT.objUserControl.Controls("Superior")
End Property

Public Sub Cargo_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Cargo_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Property Get GridComissao(ByVal objCT As Object) As Object
     Set GridComissao = objCT.objUserControl.Controls("GridComissao")
End Property

Public Property Get GridRedComissao(ByVal objCT As Object) As Object
     Set GridRedComissao = objCT.objUserControl.Controls("GridRedComissao")
End Property

Public Property Get GridRegiaoVenda(ByVal objCT As Object) As Object
     Set GridRegiaoVenda = objCT.objUserControl.Controls("GridRegiaoVenda")
End Property

Public Property Get ComissValorDe(ByVal objCT As Object) As Object
     Set ComissValorDe = objCT.objUserControl.Controls("ComissValorDe")
End Property

Public Property Get ComissValorAte(ByVal objCT As Object) As Object
     Set ComissValorAte = objCT.objUserControl.Controls("ComissValorAte")
End Property

Public Property Get ComissMoeda(ByVal objCT As Object) As Object
     Set ComissMoeda = objCT.objUserControl.Controls("ComissMoeda")
End Property

Public Property Get PercComiss(ByVal objCT As Object) As Object
     Set PercComiss = objCT.objUserControl.Controls("PercComiss")
End Property

Public Property Get RedValorDe(ByVal objCT As Object) As Object
     Set RedValorDe = objCT.objUserControl.Controls("RedValorDe")
End Property

Public Property Get RedValorAte(ByVal objCT As Object) As Object
     Set RedValorAte = objCT.objUserControl.Controls("RedValorAte")
End Property

Public Property Get RedMoeda(ByVal objCT As Object) As Object
     Set RedMoeda = objCT.objUserControl.Controls("RedMoeda")
End Property

Public Property Get PercRed(ByVal objCT As Object) As Object
     Set PercRed = objCT.objUserControl.Controls("PercRed")
End Property

Public Property Get CodRegiao(ByVal objCT As Object) As Object
     Set CodRegiao = objCT.objUserControl.Controls("CodRegiao")
End Property

Public Property Get PercComissReg(ByVal objCT As Object) As Object
     Set PercComissReg = objCT.objUserControl.Controls("PercComissReg")
End Property

Public Property Get DescRegiao(ByVal objCT As Object) As Object
     Set DescRegiao = objCT.objUserControl.Controls("DescRegiao")
End Property

Public Property Get Moeda(ByVal objCT As Object) As Object
     Set Moeda = objCT.objUserControl.Controls("Moeda")
End Property

Public Property Get ValorAporte(ByVal objCT As Object) As Object
     Set ValorAporte = objCT.objUserControl.Controls("ValorAporte")
End Property

Public Property Get PercComiAporte(ByVal objCT As Object) As Object
     Set PercComiAporte = objCT.objUserControl.Controls("PercComiAporte")
End Property

Public Sub Cargo_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Cargo_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_CARGO_VENDEDOR, Cargo(objCT), "AVISO_CRIAR_CARGO_VENDEDOR")
    If lErro <> SUCESSO Then gError 195867
    
    Exit Sub

Erro_Cargo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 195867
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195868)

    End Select

End Sub

Public Function VND_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_VND_Form_Load

    'Carrega a combo Tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_CARGO_VENDEDOR, Cargo(objCT))
    If lErro <> SUCESSO Then gError 195887

    Set objCT.gobjinfousu.objGridComissao = New AdmGrid
    Set objCT.gobjinfousu.objGridRedComissao = New AdmGrid
    Set objCT.gobjinfousu.objGridRegiaoVenda = New AdmGrid
    
    Set objCT.gobjinfousu.gobjUsuarioWeb = New ClassTRPUsuarios

    lErro = Inicializa_Grid_Comissao(objCT, objCT.gobjinfousu.objGridComissao)
    If lErro <> SUCESSO Then gError 195895

    lErro = Inicializa_Grid_RedComissao(objCT, objCT.gobjinfousu.objGridRedComissao)
    If lErro <> SUCESSO Then gError 195896

    lErro = Inicializa_Grid_RegiaoVenda(objCT, objCT.gobjinfousu.objGridRegiaoVenda)
    If lErro <> SUCESSO Then gError 195897

    lErro = Carrega_Moeda(ComissMoeda(objCT))
    If lErro <> SUCESSO Then gError 195982

    lErro = Carrega_Moeda(RedMoeda(objCT))
    If lErro <> SUCESSO Then gError 195983
    
    lErro = Carrega_Moeda(Moeda(objCT))
    If lErro <> SUCESSO Then gError 195983
    
    Moeda(objCT).ListIndex = 0
    
    objCT.PercComissaoEmissao.Text = 100
    Call objCT.PercComissaoEmissao_Validate(bSGECancelDummy)
    
    VND_Form_Load = SUCESSO
     
    Exit Function
    
Erro_VND_Form_Load:

    VND_Form_Load = gErr
     
    Select Case gErr
    
        Case 195887, 195895, 195896, 195897, 195982, 195983

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195888)
     
    End Select
     
    Exit Function

End Function

Private Function Inicializa_Grid_Comissao(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Valor De")
    objGridInt.colColuna.Add ("Valor Até")
    objGridInt.colColuna.Add ("Moeda")
    objGridInt.colColuna.Add ("% Comissão")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ComissValorDe(objCT).Name)
    objGridInt.colCampo.Add (ComissValorAte(objCT).Name)
    objGridInt.colCampo.Add (ComissMoeda(objCT).Name)
    objGridInt.colCampo.Add (PercComiss(objCT).Name)

    'Colunas do Grid
    objCT.gobjinfousu.iGrid_ComissValorDe_Col = 1
    objCT.gobjinfousu.iGrid_ComissValorAte_Col = 2
    objCT.gobjinfousu.iGrid_ComissMoeda_Col = 3
    objCT.gobjinfousu.iGrid_PercComiss_Col = 4

    'Grid do GridInterno
    objGridInt.objGrid = GridComissao(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 51

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridComissao(objCT).ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Comissao = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_RedComissao(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Valor De")
    objGridInt.colColuna.Add ("Valor Até")
    objGridInt.colColuna.Add ("Moeda")
    objGridInt.colColuna.Add ("%Máximo")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (RedValorDe(objCT).Name)
    objGridInt.colCampo.Add (RedValorAte(objCT).Name)
    objGridInt.colCampo.Add (RedMoeda(objCT).Name)
    objGridInt.colCampo.Add (PercRed(objCT).Name)

    'Colunas do Grid
    objCT.gobjinfousu.iGrid_RedValorDe_Col = 1
    objCT.gobjinfousu.iGrid_RedValorAte_Col = 2
    objCT.gobjinfousu.iGrid_RedMoeda_Col = 3
    objCT.gobjinfousu.iGrid_PercRed_Col = 4

    'Grid do GridInterno
    objGridInt.objGrid = GridRedComissao(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 51

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridRedComissao(objCT).ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_RedComissao = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_RegiaoVenda(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Região")
    objGridInt.colColuna.Add ("Descrição")
    'objGridInt.colColuna.Add ("% Dif.")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (CodRegiao(objCT).Name)
    objGridInt.colCampo.Add (DescRegiao(objCT).Name)
    'objGridInt.colCampo.Add (PercComissReg(objCT).Name)
    
    PercComissReg(objCT).left = POSICAO_FORA_TELA

    'Colunas do Grid
    objCT.gobjinfousu.iGrid_CodRegiao_Col = 1
    objCT.gobjinfousu.iGrid_DescRegiao_Col = 2
    objCT.gobjinfousu.iGrid_PercComissReg_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridRegiaoVenda(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 51

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridRegiaoVenda(objCT).ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_RegiaoVenda = SUCESSO

    Exit Function

End Function

Public Function VND_Grava_Registro(ByVal objCT As Object, objVendedor As ClassVendedor) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objVendedorTRP As New ClassVendedorTRP
Dim objVendedor1 As New ClassVendedor
Dim objVNDComissaoTRP As ClassVNDComissaoTRP
Dim objVNDReducaoTRP As ClassVNDReducaoTRP
Dim objVNDRegiaoTRP As ClassVNDRegiaoTRP

On Error GoTo Erro_VND_Grava_Registro
    
    If Len(Trim(Cargo(objCT).Text)) = 0 Then gError 195870
    
    objVendedorTRP.iCargo = Codigo_Extrai(Cargo(objCT).Text)
    
    If Len(Trim(Superior(objCT).Text)) > 0 Then objVendedor1.sNomeReduzido = Superior(objCT).Text

    'Verifica se vendedor existe
    If objVendedor1.sNomeReduzido <> "" Then
        
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor1)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 195888

        'Não encontrou o superior ==> erro
        If lErro = 25008 Then gError 195889
        
        'o superior nao pode ser o proprio vendedor
        If objVendedor.iCodigo = objVendedor1.iCodigo Then gError 195890

        objVendedorTRP.iSuperior = objVendedor1.iCodigo

    End If
    
'    'Verifica se no grid de comissao as informacoes estao preenchidas
'    For iIndice = 1 To objCT.gobjinfousu.objGridComissao.iLinhasExistentes
'
'        'se os dois valores estiverem zerados ==> erro
'        If StrParaDbl(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ComissValorDe_Col)) = 0 And StrParaDbl(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ComissValorAte_Col)) = 0 Then gError 195992
'
'        'se o valor De > valor Ate ==> erro
'        If StrParaDbl(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ComissValorDe_Col)) > StrParaDbl(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ComissValorAte_Col)) And StrParaDbl(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ComissValorAte_Col)) > 0 Then gError 195993
'
'        'se o percentual de comissao for zero ==> erro
'        If PercentParaDbl(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_PercComiss_Col)) = 0 Then gError 195994
'
'        'se a moeda nao estiver preenchida
'        If Len(Trim(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ComissMoeda_Col))) = 0 Then gError 195995
'
        Set objVNDComissaoTRP = New ClassVNDComissaoTRP
        
'        objVNDComissaoTRP.dValorDe = StrParaDbl(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ComissValorDe_Col))
'        objVNDComissaoTRP.dValorAte = StrParaDbl(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ComissValorAte_Col))
'        objVNDComissaoTRP.iMoeda = Codigo_Extrai(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ComissMoeda_Col))
'        objVNDComissaoTRP.dPercComissao = PercentParaDbl(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_PercComiss_Col))
                 
        objVNDComissaoTRP.dValorDe = 0
        objVNDComissaoTRP.dValorAte = 0
        objVNDComissaoTRP.iMoeda = 0
        objVNDComissaoTRP.dPercComissao = PercentParaDbl(objCT.perccomissao.Text)
                 
        objVendedorTRP.colVNDComissaoTRP.Add objVNDComissaoTRP
         
'    Next

'    'Verifica se no grid de comissao as informacoes estao preenchidas
'    For iIndice = 1 To objCT.gobjinfousu.objGridRedComissao.iLinhasExistentes
'
'        'se os dois valores estiverem zerados ==> erro
'        If StrParaDbl(GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_RedValorDe_Col)) = 0 And StrParaDbl(GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_RedValorAte_Col)) = 0 Then gError 195996
'
'        'se o valor De > valor Ate ==> erro
'        If StrParaDbl(GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_RedValorDe_Col)) > StrParaDbl(GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_RedValorAte_Col)) And StrParaDbl(GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_RedValorAte_Col)) > 0 Then gError 195997
'
'        'se o percentual de comissao for zero ==> erro
'        If PercentParaDbl(GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_PercRed_Col)) = 0 Then gError 195998
'
'        'se a moeda nao estiver preenchida
'        If Len(Trim(GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_RedMoeda_Col))) = 0 Then gError 195999
'
        Set objVNDReducaoTRP = New ClassVNDReducaoTRP
        
'        objVNDReducaoTRP.dValorDe = StrParaDbl(GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_RedValorDe_Col))
'        objVNDReducaoTRP.dValorAte = StrParaDbl(GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_RedValorAte_Col))
'        objVNDReducaoTRP.iMoeda = Codigo_Extrai(GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_RedMoeda_Col))
'        objVNDReducaoTRP.dPercComissaoMax = PercentParaDbl(GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_PercRed_Col))
                 
        objVNDReducaoTRP.dValorDe = StrParaDbl(ValorAporte(objCT).Text)
        objVNDReducaoTRP.dValorAte = 0
        objVNDReducaoTRP.iMoeda = Codigo_Extrai(Moeda(objCT).Text)
        objVNDReducaoTRP.dPercComissaoMax = StrParaDbl(PercComiAporte(objCT).Text) / 100
        
                 
        objVendedorTRP.colVNDReducaoTRP.Add objVNDReducaoTRP
         
'    Next

    'Verifica se no grid de comissao as informacoes estao preenchidas
    For iIndice = 1 To objCT.gobjinfousu.objGridRegiaoVenda.iLinhasExistentes
        
        Set objVNDRegiaoTRP = New ClassVNDRegiaoTRP
        
        objVNDRegiaoTRP.iRegiaoVenda = StrParaDbl(GridRegiaoVenda(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_CodRegiao_Col))
        'objVNDRegiaoTRP.dPercComissao = PercentParaDbl(GridRegiaoVenda(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_PercComissReg_Col))
                 
        objVendedorTRP.colVNDRegiaoTRP.Add objVNDRegiaoTRP
         
    Next

    Set objVendedor.objInfoUsu = objVendedorTRP
    
    Set objVendedorTRP.objUsuarioWeb = objCT.gobjinfousu.gobjUsuarioWeb

    VND_Grava_Registro = SUCESSO
    
    Exit Function
    
Erro_VND_Grava_Registro:
    
    VND_Grava_Registro = gErr
    
    Select Case gErr
    
        Case 195870
            Call Rotina_Erro(vbOKOnly, "ERRO_CARGO_NAO_INFORMADO", gErr)
        
        Case 195888
        
        Case 195889
            Call Rotina_Erro(vbOKOnly, "ERRO_SUPERIOR_NAO_CADASTRADO", gErr)
        
        Case 195890
            Call Rotina_Erro(vbOKOnly, "ERRO_SUPERIOR_IGUAL_VENDEDOR", gErr)
        
        Case 195992
            Call Rotina_Erro(vbOKOnly, "ERRO_FAIXA_VALOR_ZERADA_GRIDCOMISSAO", gErr, iIndice)

        Case 195993
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DE_MAIOR_ATE_GRIDCOMISSAO", gErr, iIndice)
        
        Case 195994
            Call Rotina_Erro(vbOKOnly, "ERRO_PERC_NAO_PREENCHIDO_GRIDCOMISSAO", gErr, iIndice)
        
        Case 195995
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_INFORMADA_GRIDCOMISSAO", gErr, iIndice)
        
        Case 195996
            Call Rotina_Erro(vbOKOnly, "ERRO_FAIXA_VALOR_ZERADA_GRIDREDUCAO", gErr, iIndice)

        Case 195997
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DE_MAIOR_ATE_GRIDREDUCAO", gErr, iIndice)
        
        Case 195998
            Call Rotina_Erro(vbOKOnly, "ERRO_PERC_NAO_PREENCHIDO_GRIDREDUCAO", gErr, iIndice)
        
        Case 195999
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_INFORMADA_GRIDREDUCAO", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195871)

    End Select

    Exit Function

End Function

Public Function VND_Traz_Vendedor_Tela(ByVal objCT As Object, ByVal objVendedor As ClassVendedor) As Long

Dim lErro As Long
Dim objVendedorTRP As ClassVendedorTRP
Dim iCodigo As Integer
Dim bCancel As Boolean
Dim objVNDComissaoTRP As ClassVNDComissaoTRP
Dim objVNDReducaoTRP As ClassVNDReducaoTRP
Dim objVNDRegiaoTRP As ClassVNDRegiaoTRP
Dim iIndice As Integer
Dim objRegiaoVenda As New ClassRegiaoVenda


On Error GoTo Erro_VND_Traz_Vendedor_Tela

    lErro = CF("Vendedor_Le_Customizado", objVendedor)
    If lErro <> SUCESSO Then gError 195877
    
    If Not (objVendedor.objInfoUsu Is Nothing) Then
   
        Set objVendedorTRP = objVendedor.objInfoUsu
        
        Set objCT.gobjinfousu.gobjUsuarioWeb = objVendedorTRP.objUsuarioWeb
        
        'Coloca Tipo no Text
        Cargo(objCT).Text = CStr(objVendedorTRP.iCargo)

        'Tenta selecionar
        lErro = Combo_Seleciona(Cargo(objCT), iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 195878
        
        If objVendedorTRP.iSuperior = 0 Then
            Superior(objCT).Text = ""
        Else
            Superior(objCT).Text = CStr(objVendedorTRP.iSuperior)
            Call Superior_Validate(objCT, bCancel)
        End If
        
        iIndice = 0
        
        Call Grid_Limpa(objCT.gobjinfousu.objGridComissao)
        
        For Each objVNDComissaoTRP In objVendedorTRP.colVNDComissaoTRP
        
            iIndice = iIndice + 1
        
            If objVNDComissaoTRP.dValorDe <> 0 Then GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ComissValorDe_Col) = Format(objVNDComissaoTRP.dValorDe, "Standard")
            If objVNDComissaoTRP.dValorAte <> 0 Then GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ComissValorAte_Col) = Format(objVNDComissaoTRP.dValorAte, "Standard")
            GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_PercComiss_Col) = Format(objVNDComissaoTRP.dPercComissao, "Percent")
            
            ComissMoeda(objCT).Text = objVNDComissaoTRP.iMoeda
            
            'Seleciona o Tipo de Cobrança
            lErro = Combo_Item_Seleciona(ComissMoeda(objCT))
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 195976
            
            GridComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ComissMoeda_Col) = ComissMoeda(objCT).Text

        Next
        
        objCT.gobjinfousu.objGridComissao.iLinhasExistentes = objVendedorTRP.colVNDComissaoTRP.Count
        
        iIndice = 0
        
        Call Grid_Limpa(objCT.gobjinfousu.objGridRedComissao)
        
        For Each objVNDReducaoTRP In objVendedorTRP.colVNDReducaoTRP
        
            iIndice = iIndice + 1
            
            If iIndice = 1 Then
                If objVNDReducaoTRP.dValorDe <> 0 Then
                    ValorAporte(objCT).Text = Format(objVNDReducaoTRP.dValorDe, "Standard")
                Else
                    ValorAporte(objCT).Text = ""
                End If
                lErro = Combo_Seleciona(Moeda(objCT), objVNDReducaoTRP.iMoeda)
                If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 195977
                
                If objVNDReducaoTRP.dPercComissaoMax <> 0 Then
                    PercComiAporte(objCT).Text = objVNDReducaoTRP.dPercComissaoMax * 100
                    Call PercComiAporte_Validate(objCT, bSGECancelDummy)
                Else
                    PercComiAporte(objCT).Text = ""
                End If
            End If
        
            If objVNDReducaoTRP.dValorDe <> 0 Then GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_RedValorDe_Col) = Format(objVNDReducaoTRP.dValorDe, "Standard")
            If objVNDReducaoTRP.dValorAte <> 0 Then GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_RedValorAte_Col) = Format(objVNDReducaoTRP.dValorAte, "Standard")
            GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_PercRed_Col) = Format(objVNDReducaoTRP.dPercComissaoMax, "Percent")
            
            RedMoeda(objCT).Text = objVNDReducaoTRP.iMoeda

            'Seleciona o Tipo de Cobrança
            lErro = Combo_Seleciona(RedMoeda(objCT), objVNDReducaoTRP.iMoeda)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 195977
            
            GridRedComissao(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_RedMoeda_Col) = RedMoeda(objCT).Text

        Next
        
        objCT.gobjinfousu.objGridRedComissao.iLinhasExistentes = objVendedorTRP.colVNDReducaoTRP.Count
        
        iIndice = 0
        
        Call Grid_Limpa(objCT.gobjinfousu.objGridRegiaoVenda)
        
        For Each objVNDRegiaoTRP In objVendedorTRP.colVNDRegiaoTRP
        
            iIndice = iIndice + 1
        
            GridRegiaoVenda(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_CodRegiao_Col) = objVNDRegiaoTRP.iRegiaoVenda
            
            objRegiaoVenda.iCodigo = objVNDRegiaoTRP.iRegiaoVenda
    
            'Tenta ler Regiao de Venda com esse código no BD
            lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
            If lErro <> SUCESSO And lErro <> 16137 Then gError 195978
            
            GridRegiaoVenda(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_DescRegiao_Col) = objRegiaoVenda.sDescricao
            
'            If objVNDRegiaoTRP.dPercComissao <> 0 Then
'                GridRegiaoVenda(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_PercComissReg_Col) = Format(objVNDRegiaoTRP.dPercComissao, "PERCENT")
'            Else
'                GridRegiaoVenda(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_PercComissReg_Col) = ""
'            End If
            
        Next
        
        objCT.gobjinfousu.objGridRegiaoVenda.iLinhasExistentes = objVendedorTRP.colVNDRegiaoTRP.Count
        
    End If
    
    VND_Traz_Vendedor_Tela = SUCESSO
     
    Exit Function
    
Erro_VND_Traz_Vendedor_Tela:

    VND_Traz_Vendedor_Tela = gErr
     
    Select Case gErr
    
        Case 195877, 195878, 195976, 195977, 195978
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195879)
     
    End Select
     
    Exit Function

End Function

Public Function VND_Limpa_Tela_Vendedor(ByVal objCT As Object) As Long
    
    Cargo(objCT).ListIndex = -1
    
    Call Grid_Limpa(objCT.gobjinfousu.objGridComissao)
    Call Grid_Limpa(objCT.gobjinfousu.objGridRedComissao)
    Call Grid_Limpa(objCT.gobjinfousu.objGridRegiaoVenda)
    
    objCT.PercComissaoEmissao.Text = 100
    Call objCT.PercComissaoEmissao_Validate(bSGECancelDummy)

    Set objCT.gobjinfousu.gobjUsuarioWeb = New ClassTRPUsuarios

    VND_Limpa_Tela_Vendedor = SUCESSO
    
End Function

Public Sub Superior_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Superior_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Superior_Validate

    If Len(Trim(Superior(objCT).Text)) > 0 Then

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le(Superior(objCT), objVendedor)
        If lErro <> SUCESSO Then gError 195884
        
        If objVendedor.iAtivo = DESMARCADO Then gError 195885
        
    End If

    Exit Sub

Erro_Superior_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 195884
        
        Case 195885
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INATIVO", gErr, objVendedor.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195886)
    
    End Select

End Sub

Public Sub LabelSuperior_Click(ByVal objCT As Object)

'BROWSE VENDEDOR :

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Superior(objCT).Text)) > 0 Then objVendedor.sNomeReduzido = Superior(objCT).Text
    
    Set gobjCT = objCT
    
    Set objEventoSuperior = New AdmEvento
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoSuperior)

End Sub

Private Sub objEventoSuperior_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1

    'Preenche campo Vendedor
    Superior(gobjCT).Text = objVendedor.sNomeReduzido

    gobjCT.iAlterado = 0

    gobjCT.objUserControl.Show

    Exit Sub

End Sub

Public Sub UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If objCT.objUserControl.ActiveControl Is Superior(objCT) Then
            Call LabelSuperior_Click(objCT)
        ElseIf objCT.objUserControl.ActiveControl Is CodRegiao(objCT) Then
            Call BotaoRegiaoVenda_Click(objCT)
        End If
    
    End If

End Sub

Public Function VND_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_VND_Saida_Celula

    'Verifica qual o Grid em questão
    Select Case objGridInt.objGrid.Name

        'Se for o GridComissao
        
        Case GridComissao(objCT).Name

            lErro = Saida_Celula_GridComissao(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195901

        'Se for o Grid de Reducao
        Case GridRedComissao(objCT).Name

            lErro = Saida_Celula_GridRedComissao(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195902

        'Se for o Grid de Regiao de Venda
        Case GridRegiaoVenda(objCT).Name

            lErro = Saida_Celula_GridRegiaoVenda(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195903

    End Select
    
    VND_Saida_Celula = SUCESSO

    Exit Function

Erro_VND_Saida_Celula:

    VND_Saida_Celula = gErr

    Select Case gErr

        Case 195901, 195902, 195903

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195904)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridComissao(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridComissao

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Produto
        Case objCT.gobjinfousu.iGrid_ComissValorDe_Col
            lErro = Saida_Celula_ComissValorDe(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195905

        Case objCT.gobjinfousu.iGrid_ComissValorAte_Col
            lErro = Saida_Celula_ComissValorAte(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195906

        Case objCT.gobjinfousu.iGrid_ComissMoeda_Col
            lErro = Saida_Celula_ComissMoeda(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195907

        'Se for a de Unidade de Medida
        Case objCT.gobjinfousu.iGrid_PercComiss_Col
            lErro = Saida_Celula_PercComiss(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195908

    End Select

    Saida_Celula_GridComissao = SUCESSO

    Exit Function

Erro_Saida_Celula_GridComissao:

    Saida_Celula_GridComissao = gErr

    Select Case gErr

        Case 195905 To 195908

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195909)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridRedComissao(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridRedComissao

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Produto
        Case objCT.gobjinfousu.iGrid_RedValorDe_Col
            lErro = Saida_Celula_RedValorDe(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195910

        Case objCT.gobjinfousu.iGrid_RedValorAte_Col
            lErro = Saida_Celula_RedValorAte(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195911

        Case objCT.gobjinfousu.iGrid_RedMoeda_Col
            lErro = Saida_Celula_RedMoeda(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195912

        'Se for a de Unidade de Medida
        Case objCT.gobjinfousu.iGrid_PercRed_Col
            lErro = Saida_Celula_PercRed(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195913

    End Select

    Saida_Celula_GridRedComissao = SUCESSO

    Exit Function

Erro_Saida_Celula_GridRedComissao:

    Saida_Celula_GridRedComissao = gErr

    Select Case gErr

        Case 195910 To 195913

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195914)

    End Select

    Exit Function

End Function


Private Function Saida_Celula_ComissValorDe(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ComissValorDe

    Set objGridInt.objControle = ComissValorDe(objCT)
    
    'Se Desconto está preenchido
    If Len(Trim(ComissValorDe(objCT).Text)) <> 0 Then

        'Verifica se o Desconto é válido
        lErro = Valor_NaoNegativo_Critica(ComissValorDe(objCT).Text)
        If lErro <> SUCESSO Then gError 195918
        
        'Acrescenta uma linha no Grid se for o caso
        If GridComissao(objCT).Row - GridComissao(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195919

    Saida_Celula_ComissValorDe = SUCESSO

    Exit Function

Erro_Saida_Celula_ComissValorDe:

    Saida_Celula_ComissValorDe = gErr

    Select Case gErr

        Case 195918, 195919
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195920)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ComissValorAte(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ComissValorAte

    Set objGridInt.objControle = ComissValorAte(objCT)
    
    'Se Desconto está preenchido
    If Len(Trim(ComissValorAte(objCT).Text)) <> 0 Then

        'Verifica se o Desconto é válido
        lErro = Valor_NaoNegativo_Critica(ComissValorAte(objCT).Text)
        If lErro <> SUCESSO Then gError 195921
        
        'Acrescenta uma linha no Grid se for o caso
        If GridComissao(objCT).Row - GridComissao(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195922

    Saida_Celula_ComissValorAte = SUCESSO

    Exit Function

Erro_Saida_Celula_ComissValorAte:

    Saida_Celula_ComissValorAte = gErr

    Select Case gErr

        Case 195921, 195922
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195923)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ComissMoeda(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_ComissMoeda

    Set objGridInt.objControle = ComissMoeda(objCT)

    'Verifica se o ComissMoeda foi preenchido
    If Len(Trim(ComissMoeda(objCT).Text)) > 0 Then

        'Verifica se ele foi selecionado
        If ComissMoeda(objCT).Text <> ComissMoeda(objCT).List(ComissMoeda(objCT).ListIndex) Then

            iCodigo = Codigo_Extrai(ComissMoeda(objCT).Text)

            'Seleciona o Tipo de Cobrança
            lErro = Combo_Seleciona(ComissMoeda(objCT), iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 195924

            If lErro = 6730 Then gError 195925
            If lErro = 6731 Then gError 195926

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissao(objCT).Row - GridComissao(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195927

    Saida_Celula_ComissMoeda = SUCESSO

    Exit Function

Erro_Saida_Celula_ComissMoeda:

    Saida_Celula_ComissMoeda = gErr

    Select Case gErr

        Case 195924, 195927
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 195925, 195926
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_CADASTRADA", gErr, ComissMoeda(objCT).Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195928)

    End Select

    Exit Function

End Function

Function Saida_Celula_PercComiss(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double

On Error GoTo Erro_Saida_Celula_PercComiss

    Set objGridInt.objControle = PercComiss(objCT)

    If Len(PercComiss(objCT).Text) > 0 Then
        
        'Critica a porcentagem
        lErro = Porcentagem_Critica(PercComiss(objCT).Text)
        If lErro <> SUCESSO Then gError 195929

        dPercentDesc = CDbl(PercComiss(objCT).Text)
        
        PercComiss(objCT).Text = Format(dPercentDesc, "Fixed")
        
        If GridComissao(objCT).Row - GridComissao(objCT).FixedRows = objGridInt.iLinhasExistentes Then
    
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195930

    Saida_Celula_PercComiss = SUCESSO

    Exit Function

Erro_Saida_Celula_PercComiss:

    Saida_Celula_PercComiss = gErr

    Select Case gErr

        Case 195929, 195930
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195931)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_RedValorDe(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_RedValorDe

    Set objGridInt.objControle = RedValorDe(objCT)
    
    'Se Desconto está preenchido
    If Len(Trim(RedValorDe(objCT).Text)) <> 0 Then

        lErro = Valor_NaoNegativo_Critica(RedValorDe(objCT).Text)
        If lErro <> SUCESSO Then gError 195932
        
        'Acrescenta uma linha no Grid se for o caso
        If GridRedComissao(objCT).Row - GridRedComissao(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195933

    Saida_Celula_RedValorDe = SUCESSO

    Exit Function

Erro_Saida_Celula_RedValorDe:

    Saida_Celula_RedValorDe = gErr

    Select Case gErr

        Case 195932, 195933
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195934)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_RedValorAte(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_RedValorAte

    Set objGridInt.objControle = RedValorAte(objCT)
    
    'Se Desconto está preenchido
    If Len(Trim(RedValorAte(objCT).Text)) <> 0 Then

        lErro = Valor_NaoNegativo_Critica(RedValorAte(objCT).Text)
        If lErro <> SUCESSO Then gError 195935
        
        'Acrescenta uma linha no Grid se for o caso
        If GridRedComissao(objCT).Row - GridRedComissao(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195936

    Saida_Celula_RedValorAte = SUCESSO

    Exit Function

Erro_Saida_Celula_RedValorAte:

    Saida_Celula_RedValorAte = gErr

    Select Case gErr

        Case 195935, 195936
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195937)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_RedMoeda(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_RedMoeda

    Set objGridInt.objControle = RedMoeda(objCT)

    'Verifica se o RedMoeda foi preenchido
    If Len(Trim(RedMoeda(objCT).Text)) > 0 Then

        'Verifica se ele foi selecionado
        If RedMoeda(objCT).Text <> RedMoeda(objCT).List(RedMoeda(objCT).ListIndex) Then

            'Seleciona o Tipo de Cobrança
            lErro = Combo_Seleciona(RedMoeda(objCT), iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 195938

            If lErro = 6730 Then gError 195939
            If lErro = 6731 Then gError 195940

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridRedComissao(objCT).Row - GridRedComissao(objCT).FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195941

    Saida_Celula_RedMoeda = SUCESSO

    Exit Function

Erro_Saida_Celula_RedMoeda:

    Saida_Celula_RedMoeda = gErr

    Select Case gErr

        Case 195938, 195941
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 195939, 195940
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_CADASTRADA", gErr, RedMoeda(objCT).Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195942)

    End Select

    Exit Function

End Function

Function Saida_Celula_PercRed(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual de Reducao que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double

On Error GoTo Erro_Saida_Celula_PercRed

    Set objGridInt.objControle = PercRed(objCT)

    If Len(PercRed(objCT).Text) > 0 Then
        
        'Critica a porcentagem
        lErro = Porcentagem_Critica(PercRed(objCT).Text)
        If lErro <> SUCESSO Then gError 195943

        dPercentDesc = CDbl(PercRed(objCT).Text)
        
        PercRed(objCT).Text = Format(dPercentDesc, "Fixed")
        
        If GridRedComissao(objCT).Row - GridRedComissao(objCT).FixedRows = objGridInt.iLinhasExistentes Then
    
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195944

    Saida_Celula_PercRed = SUCESSO

    Exit Function

Erro_Saida_Celula_PercRed:

    Saida_Celula_PercRed = gErr

    Select Case gErr

        Case 195943, 195944
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195945)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CodRegiao(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_CodRegiao

    Set objGridInt.objControle = CodRegiao(objCT)

    'Verifica se foi preenchido o campo Regiao
    If Len(Trim(CodRegiao(objCT).Text)) <> 0 Then

        'Verifica se já está em outra linha do Grid
        For iIndice = 1 To objCT.gobjinfousu.objGridRegiaoVenda.iLinhasExistentes
            If iIndice <> GridRegiaoVenda(objCT).Row Then
                If GridRegiaoVenda(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_CodRegiao_Col) = CodRegiao(objCT).Text Then gError 195986
            End If
        Next

        'Verifica se está preenchida com o ítem selecionado na ComboBox Regiao
        If CodRegiao(objCT).Text <> GridRegiaoVenda(objCT).TextMatrix(GridRegiaoVenda(objCT).Row, objCT.gobjinfousu.iGrid_CodRegiao_Col) Then

            objRegiaoVenda.iCodigo = StrParaInt(CodRegiao(objCT).Text)
    
            'Tenta ler Regiao de Venda com esse código no BD
            lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
            If lErro <> SUCESSO And lErro <> 16137 Then gError 195946
            
            'Não encontrou Regiao Venda BD
            If lErro <> SUCESSO Then gError 195947
            
            GridRegiaoVenda(objCT).TextMatrix(GridRegiaoVenda(objCT).Row, objCT.gobjinfousu.iGrid_DescRegiao_Col) = objRegiaoVenda.sDescricao
    
        End If

        If GridRegiaoVenda(objCT).Row - GridRegiaoVenda(objCT).FixedRows = objGridInt.iLinhasExistentes Then
    
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195944

    Saida_Celula_CodRegiao = SUCESSO
    
    Exit Function

Erro_Saida_Celula_CodRegiao:
    
    Saida_Celula_CodRegiao = gErr
    
    Select Case gErr

        Case 195946
    
        Case 195947
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_REGIAO")
    
            If vbMsgRes = vbYes Then
                'Chama a tela RegiaoVenda
                Call Chama_Tela("RegiaoVenda", objRegiaoVenda)
    
            End If
    
        Case 195986
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_JA_PREENCHIDO_LINHA_GRID", gErr, iIndice)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 195948)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridRegiaoVenda(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridRegiaoVenda

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Produto
        Case objCT.gobjinfousu.iGrid_CodRegiao_Col
            lErro = Saida_Celula_CodRegiao(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195949

        Case objCT.gobjinfousu.iGrid_PercComissReg_Col
            lErro = Saida_Celula_PercComissReg(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195949

    End Select

    Saida_Celula_GridRegiaoVenda = SUCESSO

    Exit Function

Erro_Saida_Celula_GridRegiaoVenda:

    Saida_Celula_GridRegiaoVenda = gErr

    Select Case gErr

        Case 195949

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195950)

    End Select

    Exit Function

End Function

Public Sub GridComissao_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjinfousu.objGridComissao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjinfousu.objGridComissao, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridComissao_EnterCell(ByVal objCT As Object)
    Call Grid_Entrada_Celula(objCT.gobjinfousu.objGridComissao, objCT.iAlterado)
End Sub

Public Sub GridComissao_GotFocus(ByVal objCT As Object)
    Call Grid_Recebe_Foco(objCT.gobjinfousu.objGridComissao)
End Sub

Public Sub GridComissao_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjinfousu.objGridComissao)
    
End Sub

Public Sub GridComissao_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjinfousu.objGridComissao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjinfousu.objGridComissao, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridComissao_LeaveCell(ByVal objCT As Object)
    Call objCT.Saida_Celula(objCT.gobjinfousu.objGridComissao)
End Sub

Public Sub GridComissao_Validate(ByVal objCT As Object, Cancel As Boolean)
    Call Grid_Libera_Foco(objCT.gobjinfousu.objGridComissao)
End Sub

Public Sub GridComissao_RowColChange(ByVal objCT As Object)
    Call Grid_RowColChange(objCT.gobjinfousu.objGridComissao)
End Sub

Public Sub GridComissao_Scroll(ByVal objCT As Object)
    Call Grid_Scroll(objCT.gobjinfousu.objGridComissao)
End Sub

Public Sub GridRedComissao_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjinfousu.objGridRedComissao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjinfousu.objGridRedComissao, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridRedComissao_EnterCell(ByVal objCT As Object)
    Call Grid_Entrada_Celula(objCT.gobjinfousu.objGridRedComissao, objCT.iAlterado)
End Sub

Public Sub GridRedComissao_GotFocus(ByVal objCT As Object)
    Call Grid_Recebe_Foco(objCT.gobjinfousu.objGridRedComissao)
End Sub

Public Sub GridRedComissao_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjinfousu.objGridRedComissao)
    
End Sub

Public Sub GridRedComissao_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjinfousu.objGridRedComissao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjinfousu.objGridRedComissao, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridRedComissao_LeaveCell(ByVal objCT As Object)
    Call objCT.Saida_Celula(objCT.gobjinfousu.objGridRedComissao)
End Sub

Public Sub GridRedComissao_Validate(ByVal objCT As Object, Cancel As Boolean)
    Call Grid_Libera_Foco(objCT.gobjinfousu.objGridRedComissao)
End Sub

Public Sub GridRedComissao_RowColChange(ByVal objCT As Object)
    Call Grid_RowColChange(objCT.gobjinfousu.objGridRedComissao)
End Sub

Public Sub GridRedComissao_Scroll(ByVal objCT As Object)
    Call Grid_Scroll(objCT.gobjinfousu.objGridRedComissao)
End Sub

Public Sub GridRegiaoVenda_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjinfousu.objGridRegiaoVenda, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjinfousu.objGridRegiaoVenda, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridRegiaoVenda_EnterCell(ByVal objCT As Object)
    Call Grid_Entrada_Celula(objCT.gobjinfousu.objGridRegiaoVenda, objCT.iAlterado)
End Sub

Public Sub GridRegiaoVenda_GotFocus(ByVal objCT As Object)
    Call Grid_Recebe_Foco(objCT.gobjinfousu.objGridRegiaoVenda)
End Sub

Public Sub GridRegiaoVenda_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjinfousu.objGridRegiaoVenda)
    
End Sub

Public Sub GridRegiaoVenda_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjinfousu.objGridRegiaoVenda, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjinfousu.objGridRegiaoVenda, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridRegiaoVenda_LeaveCell(ByVal objCT As Object)
    Call objCT.Saida_Celula(objCT.gobjinfousu.objGridRegiaoVenda)
End Sub

Public Sub GridRegiaoVenda_Validate(ByVal objCT As Object, Cancel As Boolean)
    Call Grid_Libera_Foco(objCT.gobjinfousu.objGridRegiaoVenda)
End Sub

Public Sub GridRegiaoVenda_RowColChange(ByVal objCT As Object)
    Call Grid_RowColChange(objCT.gobjinfousu.objGridRegiaoVenda)
End Sub

Public Sub GridRegiaoVenda_Scroll(ByVal objCT As Object)
    Call Grid_Scroll(objCT.gobjinfousu.objGridRegiaoVenda)
End Sub

Public Sub ComissValorDe_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ComissValorDe_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjinfousu.objGridComissao)
End Sub

Public Sub ComissValorDe_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjinfousu.objGridComissao)
End Sub

Public Sub ComissValorDe_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjinfousu.objGridComissao.objControle = ComissValorDe(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjinfousu.objGridComissao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ComissValorAte_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ComissValorAte_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjinfousu.objGridComissao)
End Sub

Public Sub ComissValorAte_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjinfousu.objGridComissao)
End Sub

Public Sub ComissValorAte_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjinfousu.objGridComissao.objControle = ComissValorAte(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjinfousu.objGridComissao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ComissMoeda_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ComissMoeda_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjinfousu.objGridComissao)
End Sub

Public Sub ComissMoeda_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjinfousu.objGridComissao)
End Sub

Public Sub ComissMoeda_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjinfousu.objGridComissao.objControle = ComissMoeda(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjinfousu.objGridComissao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub PercComiss_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercComiss_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjinfousu.objGridComissao)
End Sub

Public Sub PercComiss_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjinfousu.objGridComissao)
End Sub

Public Sub PercComiss_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjinfousu.objGridComissao.objControle = PercComiss(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjinfousu.objGridComissao)
    If lErro <> SUCESSO Then Cancel = True
    
    
End Sub

Public Sub RedValorDe_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub RedValorDe_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjinfousu.objGridRedComissao)
End Sub

Public Sub RedValorDe_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjinfousu.objGridRedComissao)
End Sub

Public Sub RedValorDe_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjinfousu.objGridRedComissao.objControle = RedValorDe(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjinfousu.objGridRedComissao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub RedValorAte_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub RedValorAte_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjinfousu.objGridRedComissao)
End Sub

Public Sub RedValorAte_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjinfousu.objGridRedComissao)
End Sub

Public Sub RedValorAte_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjinfousu.objGridRedComissao.objControle = RedValorAte(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjinfousu.objGridRedComissao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub RedMoeda_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub RedMoeda_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjinfousu.objGridRedComissao)
End Sub

Public Sub RedMoeda_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjinfousu.objGridRedComissao)
End Sub

Public Sub RedMoeda_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjinfousu.objGridRedComissao.objControle = RedMoeda(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjinfousu.objGridRedComissao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub PercRed_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercRed_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjinfousu.objGridRedComissao)
End Sub

Public Sub PercRed_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjinfousu.objGridRedComissao)
End Sub

Public Sub PercRed_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjinfousu.objGridRedComissao.objControle = PercRed(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjinfousu.objGridRedComissao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub CodRegiao_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CodRegiao_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjinfousu.objGridRegiaoVenda)
End Sub

Public Sub CodRegiao_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjinfousu.objGridRegiaoVenda)
End Sub

Public Sub CodRegiao_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjinfousu.objGridRegiaoVenda.objControle = CodRegiao(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjinfousu.objGridRegiaoVenda)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Function Carrega_Moeda(objCombo As Object) As Long

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError 195979
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 195980
    
    For Each objMoeda In colMoedas
    
        objCombo.AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
        objCombo.ItemData(objCombo.NewIndex) = objMoeda.iCodigo
    Next

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 195979
        
        Case 195980
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195981)
    
    End Select

End Function

Public Sub BotaoRegiaoVenda_Click(ByVal objCT As Object)

Dim objRegiao As New ClassRegiaoVenda
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(CodRegiao(objCT).Text)) > 0 Then objRegiao.iCodigo = StrParaInt(CodRegiao(objCT).Text)
    
    Set gobjCT = objCT
    
    Set objEventoRegiao = New AdmEvento
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("RegiaoVendaLista", colSelecao, objRegiao, objEventoRegiao)

End Sub

Private Sub objEventoRegiao_evSelecao(obj1 As Object)

Dim objRegiao As ClassRegiaoVenda
Dim iIndice As Integer

On Error GoTo Erro_objEventoRegiao_evSelecao

    'verifica se tem alguma linha do Grid selecionada
    If GridRegiaoVenda(gobjCT).Row > 0 Then

        Set objRegiao = obj1

        CodRegiao(gobjCT) = objRegiao.iCodigo

        If Not (gobjCT.objUserControl.ActiveControl Is CodRegiao(gobjCT)) Then
    
            'Verifica se já está em outra linha do Grid
            For iIndice = 1 To gobjCT.gobjinfousu.objGridRegiaoVenda.iLinhasExistentes
                If iIndice <> GridRegiaoVenda(gobjCT).Row Then
                    If GridRegiaoVenda(gobjCT).TextMatrix(iIndice, gobjCT.gobjinfousu.iGrid_CodRegiao_Col) = CodRegiao(gobjCT).Text Then gError 195984
                End If
            Next

    
            'Preenche o Grid
            GridRegiaoVenda(gobjCT).TextMatrix(GridRegiaoVenda(gobjCT).Row, gobjCT.gobjinfousu.iGrid_CodRegiao_Col) = CodRegiao(gobjCT).Text
            GridRegiaoVenda(gobjCT).TextMatrix(GridRegiaoVenda(gobjCT).Row, gobjCT.gobjinfousu.iGrid_DescRegiao_Col) = objRegiao.sDescricao
    
            If GridRegiaoVenda(gobjCT).Row - GridRegiaoVenda(gobjCT).FixedRows = gobjCT.gobjinfousu.objGridRegiaoVenda.iLinhasExistentes Then
        
                gobjCT.gobjinfousu.objGridRegiaoVenda.iLinhasExistentes = gobjCT.gobjinfousu.objGridRegiaoVenda.iLinhasExistentes + 1
        
            End If
    
    
        End If
        

    End If

    gobjCT.objUserControl.Show

    Exit Sub

Erro_objEventoRegiao_evSelecao:

    Select Case gErr

        Case 195984
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_JA_PREENCHIDO_LINHA_GRID", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195985)

    End Select

    Exit Sub

End Sub

Function Saida_Celula_PercComissReg(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual de Reducao que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double

On Error GoTo Erro_Saida_Celula_PercComissReg

    Set objGridInt.objControle = PercComissReg(objCT)

    If Len(PercComissReg(objCT).Text) > 0 Then
        
        'Critica a porcentagem
        lErro = Porcentagem_Critica(PercComissReg(objCT).Text)
        If lErro <> SUCESSO Then gError 195943

        dPercentDesc = CDbl(PercComissReg(objCT).Text)
        
        PercComissReg(objCT).Text = Format(dPercentDesc, "Fixed")
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195944

    Saida_Celula_PercComissReg = SUCESSO

    Exit Function

Erro_Saida_Celula_PercComissReg:

    Saida_Celula_PercComissReg = gErr

    Select Case gErr

        Case 195943, 195944
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195945)

    End Select

    Exit Function

End Function

Public Sub PercComissReg_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercComissReg_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjinfousu.objGridRegiaoVenda)
End Sub

Public Sub PercComissReg_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjinfousu.objGridRegiaoVenda)
End Sub

Public Sub PercComissReg_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjinfousu.objGridRegiaoVenda.objControle = PercComissReg(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjinfousu.objGridRegiaoVenda)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Moeda_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorAporte_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercComiAporte_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercComiAporte_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiAporte_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComiAporte(objCT).Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComiAporte(objCT).Text)
    If lErro <> SUCESSO Then Error 195853

    'Formata
    PercComiAporte(objCT).Text = Format(PercComiAporte(objCT).Text, "Fixed")

    Exit Sub

Erro_PercComiAporte_Validate:

    Cancel = True

    Select Case Err

        Case 195853
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 195854)
    
    End Select

    Exit Sub

End Sub

Public Sub ValorAporte_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorAporte_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(ValorAporte(objCT).Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Valor_Positivo_Critica(ValorAporte(objCT).Text)
    If lErro <> SUCESSO Then gError 195853

    'Formata
    ValorAporte(objCT).Text = Format(ValorAporte(objCT).Text, "STANDARD")

    Exit Sub

Erro_ValorAporte_Validate:

    Cancel = True

    Select Case gErr

        Case 195853
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195854)
    
    End Select

    Exit Sub

End Sub

Public Sub BotaoWebAcesso_Click(ByVal objCT As Object)
    objCT.gobjinfousu.gobjUsuarioWeb.iTipoUsuario = TRP_USUARIO_VENDEDOR
    objCT.gobjinfousu.gobjUsuarioWeb.sNome = objCT.NomeReduzido.Text
    objCT.gobjinfousu.gobjUsuarioWeb.lCodigo = StrParaDbl(objCT.Codigo.Text)
    Call Chama_Tela_Modal("TRPUsuarios", objCT.gobjinfousu.gobjUsuarioWeb)
End Sub
