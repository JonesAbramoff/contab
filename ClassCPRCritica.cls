VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCPRCritica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function TipoCliente_CarregaCombo(ByVal ComboTipo As Object) As Long
'preenche combo com codigo-descricao de tipos de cliente setando itemdata

Dim lErro As Long
Dim colCodigoDescricaoCliente As New AdmColCodigoNome
Dim objCodigoDescricaoCliente As New AdmCodigoNome

On Error GoTo Erro_TipoCliente_CarregaCombo

    lErro = CF("Cod_Nomes_Le", "TiposdeCliente", "Codigo", "Descricao", STRING_TIPO_CLIENTE_DESCRICAO, colCodigoDescricaoCliente)
    If lErro <> SUCESSO Then Error 47617
    
   'preenche a ListBox ComboTipo com os objetos da colecao
    For Each objCodigoDescricaoCliente In colCodigoDescricaoCliente
        ComboTipo.AddItem objCodigoDescricaoCliente.iCodigo & SEPARADOR & objCodigoDescricaoCliente.sNome
        ComboTipo.ItemData(ComboTipo.NewIndex) = objCodigoDescricaoCliente.iCodigo
    Next
        
    TipoCliente_CarregaCombo = SUCESSO

    Exit Function
    
Erro_TipoCliente_CarregaCombo:

    TipoCliente_CarregaCombo = Err

    Select Case Err

        Case 47617
        
        Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147893)

    End Select

    Exit Function

End Function

Function TipoCliente_ValidaCombo(ByVal ComboTipo As Object) As Long

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_TipoCliente_ValidaCombo

    'Verifica se foi preenchida a ComboBox Tipo
    If Len(Trim(ComboTipo.Text)) <> 0 Then

        'Verifica se está preenchida com o item selecionado na ComboBox Tipo
        If ComboTipo.Text <> ComboTipo.List(ComboTipo.ListIndex) Then
        
            'Verifica se existe o ítem na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(ComboTipo, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 56726
        
            'Não existe o ítem com a STRING na List da ComboBox
            If lErro <> SUCESSO Then Error 56727
        
        End If
        
    End If

    TipoCliente_ValidaCombo = SUCESSO
    
    Exit Function

Erro_TipoCliente_ValidaCombo:

    TipoCliente_ValidaCombo = Err
    
    Select Case Err

        Case 56726 'Tratado na rotina chamada
    
        Case 56727
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_ENCONTRADO2", Err, ComboTipo.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147894)

    End Select

    Exit Function

End Function

Function TipoFornecedor_ValidaCombo(ByVal ComboTipo As Object) As Long

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_TipoFornecedor_ValidaCombo

    'Verifica se foi preenchida a ComboBox Tipo
    If Len(Trim(ComboTipo.Text)) <> 0 Then

        'Verifica se está preenchida com o item selecionado na ComboBox Tipo
        If ComboTipo.Text <> ComboTipo.List(ComboTipo.ListIndex) Then
        
            'Verifica se existe o ítem na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(ComboTipo, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 56724
        
            'Não existe o ítem com a STRING na List da ComboBox
            If lErro <> SUCESSO Then Error 56725
        
        End If
        
    End If

    TipoFornecedor_ValidaCombo = SUCESSO
    
    Exit Function

Erro_TipoFornecedor_ValidaCombo:

    TipoFornecedor_ValidaCombo = Err
    
    Select Case Err

        Case 56724 'Tratado na rotina chamada
    
        Case 56725
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORNECEDOR_NAO_ENCONTRADO2", Err, ComboTipo.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147895)

    End Select

    Exit Function

End Function

Function TipoFornecedor_CarregaCombo(ByVal ComboTipo As Object) As Long

Dim lErro As Long
Dim colCodigoDescricaoFornecedor As New AdmColCodigoNome
Dim objCodigoDescricaoFornecedor As New AdmCodigoNome

On Error GoTo Erro_TipoFornecedor_CarregaCombo

    lErro = CF("Cod_Nomes_Le", "TiposdeFornecedor", "Codigo", "Descricao", STRING_TIPO_FORNECEDOR_DESCRICAO, colCodigoDescricaoFornecedor)
    If lErro <> SUCESSO Then Error 47656
    
   'preenche a ListBox ComboTipo com os objetos da colecao
    For Each objCodigoDescricaoFornecedor In colCodigoDescricaoFornecedor
        ComboTipo.AddItem objCodigoDescricaoFornecedor.iCodigo & SEPARADOR & objCodigoDescricaoFornecedor.sNome
        ComboTipo.ItemData(ComboTipo.NewIndex) = objCodigoDescricaoFornecedor.iCodigo
    Next
        
    TipoFornecedor_CarregaCombo = SUCESSO

    Exit Function
    
Erro_TipoFornecedor_CarregaCombo:

    TipoFornecedor_CarregaCombo = Err

    Select Case Err

    Case 47656
    
    Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147896)

    End Select

    Exit Function

End Function

Function ContaCorrente_Bancaria_ValidaCombo(ByVal ComboContaCorrente As Object) As Long

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim iCodigo As Integer

On Error GoTo Erro_ContaCorrente_Bancaria_Valida

    If Len(Trim(ComboContaCorrente.Text)) <> 0 Then

        'Verifica se esta preenchida com o item selecionado na ComboBox ComboContaCorrente
        If ComboContaCorrente.Text <> ComboContaCorrente.List(ComboContaCorrente.ListIndex) Then
    
            lErro = Combo_Seleciona(ComboContaCorrente, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 23268
        
            'Nao existe o ítem com o CÓDIGO na List da ComboBox
            If lErro = 6730 Then
        
                objContaCorrenteInt.iCodigo = iCodigo
                lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
                If lErro <> SUCESSO And lErro <> 11807 Then Error 23269
        
                'Não encontrou a Conta Corrente no BD
                If lErro = 11807 Then Error 23270
        
                'se a conta nao é bancaria
                If objContaCorrenteInt.iCodBanco = 0 Then Error 56735
                
                'Encontrou a Conta Corrente no BD, coloca no Text da Combo
                ComboContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido
        
            End If
        
            'Não existe o ítem com a STRING na List da ComboBox
            If lErro = 6731 Then Error 23271

        End If
    
    End If
    
    ContaCorrente_Bancaria_ValidaCombo = SUCESSO
    
    Exit Function

Erro_ContaCorrente_Bancaria_Valida:

    ContaCorrente_Bancaria_ValidaCombo = Err
    
    Select Case Err

        Case 23268, 23269

        Case 23270
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objContaCorrenteInt.iCodigo)
            
        Case 23271
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE1", Err, ComboContaCorrente.Text)
                   
        Case 56735
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", Err, ComboContaCorrente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147897)

    End Select

    Exit Function

End Function

Function ContasCorrentes_Bancarias_CarregaCombo(ByVal ComboContaCorrente As Object) As Long
'Carrega as contas correntes bancarias na combo de contas correntes

Dim lErro As Long
Dim colCodigoNomeConta As New AdmColCodigoNome
Dim objCodigoNomeConta As AdmCodigoNome

On Error GoTo Erro_ContasCorrentes_Bancarias_CarregaCombo

    'Carrega a Coleção de Contas
    lErro = CF("ContasCorrentes_Bancarias_Le_CodigosNomesRed", colCodigoNomeConta)
    If lErro <> SUCESSO Then Error 23267

    'Preenche a ComboBox CodConta com os objetos da coleção de Contas
    For Each objCodigoNomeConta In colCodigoNomeConta

        ComboContaCorrente.AddItem CStr(objCodigoNomeConta.iCodigo) & SEPARADOR & objCodigoNomeConta.sNome
        ComboContaCorrente.ItemData(ComboContaCorrente.NewIndex) = objCodigoNomeConta.iCodigo

    Next

    ContasCorrentes_Bancarias_CarregaCombo = SUCESSO

    Exit Function

Erro_ContasCorrentes_Bancarias_CarregaCombo:

    ContasCorrentes_Bancarias_CarregaCombo = Err

    Select Case Err

        Case 23267

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147898)

    End Select

    Exit Function

End Function

Public Function Transportadora_Padrao_FilCli(ByVal objFilialCliente As ClassFilialCliente, ByVal objTransportadora As ClassTransportadora) As Long
'preenche objTransportadora à partir da filial cliente ( se ela nao tiver valor default tenta pegar do tipo de cliente )

Dim lErro As Long
Dim objTipoCliente As New ClassTipoCliente
Dim objCliente As New ClassCliente

On Error GoTo Erro_Transportadora_Padrao_FilCli

    If objFilialCliente.iCodTransportadora > 0 Then
        
        objTransportadora.iCodigo = objFilialCliente.iCodTransportadora
    
    Else
    
        objCliente.lCodigo = objFilialCliente.lCodCliente
        
        'Lê o Cliente
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then Error 59024
        
        'Não achou o Cliente --> erro
        If lErro = 12293 Then Error 59025
                
        If objCliente.iTipo <> 0 Then
            
            objTipoCliente.iCodigo = objCliente.iTipo
            
            'Lê o Tipo de Cliente
            lErro = CF("TipoCliente_Le", objTipoCliente)
            If lErro <> SUCESSO And lErro <> 19062 Then Error 59026
            
            'Não achou o Tipo de Cliente --> erro
            If lErro = 19062 Then Error 59027
            
            objTransportadora.iCodigo = objTipoCliente.iCodTransportadora
    
        End If
        
    End If

    If objTransportadora.iCodigo <> 0 Then
    
        'Lê a transportadora
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 59028
        If lErro <> SUCESSO Then Error 59029
        
    End If
    
    Transportadora_Padrao_FilCli = SUCESSO
     
    Exit Function
    
Erro_Transportadora_Padrao_FilCli:

    Transportadora_Padrao_FilCli = Err
     
    Select Case Err
          
        Case 59024, 59026, 59028
        
        Case 59025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objCliente.lCodigo)
            
        Case 59027
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_CADASTRADO", Err, objTipoCliente.iCodigo)
            
        Case 59029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objTransportadora.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147899)
     
    End Select
     
    Exit Function

End Function

Function NatMovCta_Critica_NatMovCtaPai(ByVal sCodigo As String) As Long
'verifica se o codigo tem um pai. Se não tiver -> ERRO

Dim sCodigo1 As String
Dim iNivel As Integer
Dim lErro As Long
Dim objNatMovCta As New ClassNatMovCta

On Error GoTo Erro_NatMovCta_Critica_NatMovCtaPai

    'retorna o nivel do código sCodigo
    lErro = Mascara_Item_ObterNivel(SEGMENTO_NATMOVCTA, sCodigo, iNivel)
    If lErro <> SUCESSO Then gError 122800
    
    'se tem um "pai"
    If iNivel > 1 Then
    
        sCodigo1 = String(STRING_NATMOVCTA_CODIGO, 0)
    
        'descobre o "pai"
        lErro = Mascara_RetornaItemNoNivel(SEGMENTO_NATMOVCTA, iNivel - 1, sCodigo, sCodigo1)
        If lErro <> SUCESSO Then gError 122801
    
        'Obj recebe o código do "pai"
        objNatMovCta.sCodigo = sCodigo1
        
        'verifica se o "pai" está cadastrado
        lErro = CF("NatMovCta_Le", objNatMovCta)
        If lErro <> SUCESSO And lErro <> 122786 Then gError 122802
        
        'nao encontrou "pai"
        If lErro = 122786 Then gError 122803
        
    End If
        
    NatMovCta_Critica_NatMovCtaPai = SUCESSO
    
    Exit Function
    
Erro_NatMovCta_Critica_NatMovCtaPai:

    NatMovCta_Critica_NatMovCtaPai = gErr
    
    Select Case gErr
    
        Case 122800
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_ITEM_OBTERNIVEL", gErr)
    
        Case 122801
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAITEMNONIVEL", gErr, iNivel - 1)
    
        Case 122802
            
        Case 122803
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMOVTO1PAI_INEXISTENTE", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147900)
        
    End Select

    Exit Function

End Function

Function Natureza_Critica(ByVal objNatMovCta As ClassNatMovCta, Optional ByVal iTipoMov As Integer = 0) As Long
'Verifica se a natureza com o código contido em objNatMovCta é do Tipo de Movimentação passado como
'parâmetro (pagto ou recebimento), além de verificar se ele é analítico. Caso Contrário -> ERRO

Dim lErro As Long
Dim iNivel As Integer, iTemFilho As Integer

On Error GoTo Erro_Natureza_Critica

     'lê o registro referente ao código
     lErro = CF("NatMovCta_Le", objNatMovCta)
     If lErro <> SUCESSO And lErro <> 122786 Then gError 122841
        
     If lErro = 122786 Then gError 122842
     
     'Se o Tipo de Movimentação não for do tipo passado como parâmetro -> ERRO
     If iTipoMov <> 0 And objNatMovCta.iTipo <> iTipoMov Then gError 122844
     
     lErro = CF("NatMovCta_Tem_Filho", objNatMovCta.sCodigo, iTemFilho)
     If lErro <> SUCESSO Then gError 122845
     
     If iTemFilho = ITEM_TEM_FILHOS Then gError 122846
     
     Natureza_Critica = SUCESSO
     
     Exit Function
    
Erro_Natureza_Critica:

    Natureza_Critica = gErr
    
    Select Case gErr

        Case 122841, 122845
        
        Case 122842
            Call Rotina_Erro(vbOKOnly, "ERRO_NATMOVCTA_NAO CADASTRADA", gErr, objNatMovCta.sCodigo)
                    
        Case 122844
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NATMOVCTA", gErr, objNatMovCta.sCodigo)
            
        Case 122846
            Call Rotina_Erro(vbOKOnly, "ERRO_NATMOVCTA_SINTETICA", gErr, objNatMovCta.sCodigo)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147901)
    
    End Select
    
    Exit Function
    
End Function


