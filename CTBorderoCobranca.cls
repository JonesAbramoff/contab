VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTBorderoCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjInfoUsu As Object

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Private gobjBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao

'--- inicio dos properties get dos controles da tela

Public Property Get BotaoSeguir() As Object
     Set BotaoSeguir = objUserControl.Controls("BotaoSeguir")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get Cobrador() As Object
     Set Cobrador = objUserControl.Controls("Cobrador")
End Property

Public Property Get Carteira() As Object
     Set Carteira = objUserControl.Controls("Carteira")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get TipoCliente() As Object
     Set TipoCliente = objUserControl.Controls("TipoCliente")
End Property

Public Property Get Estado() As Object
     Set Estado = objUserControl.Controls("Estado")
End Property

Public Property Get UpDownDataVencimento() As Object
     Set UpDownDataVencimento = objUserControl.Controls("UpDownDataVencimento")
End Property

Public Property Get DataVencimento() As Object
     Set DataVencimento = objUserControl.Controls("DataVencimento")
End Property

Public Property Get Totalizando() As Object
     Set Totalizando = objUserControl.Controls("Totalizando")
End Property

Public Property Get UpDownDataEmissao() As Object
     Set UpDownDataEmissao = objUserControl.Controls("UpDownDataEmissao")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get UpDownDataContabil() As Object
     Set UpDownDataContabil = objUserControl.Controls("UpDownDataContabil")
End Property

Public Property Get DataContabil() As Object
     Set DataContabil = objUserControl.Controls("DataContabil")
End Property

Public Property Get LabelDataContabil() As Object
     Set LabelDataContabil = objUserControl.Controls("LabelDataContabil")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Public Sub BotaoSeguir_Click()

Dim lErro As Long
Dim iCodCobrador As Integer
Dim iTipoCobranca As Integer
Dim objCobrador As New ClassCobrador
Dim objCarteira As ClassCarteiraCobrador

On Error GoTo Erro_BotaoSeguir_Click

    'Verifica se o Cobrador está preenchido
    If Len(Trim(Cobrador.Text)) = 0 Then gError 15795

    'Verifica se a Carteira está preenchida
    If Len(Trim(Carteira.Text)) = 0 Then gError 15796

    'Verifica se a DataEmissao está preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 15797

    'Verifica se a DataVencimento está preenchida
    If Len(Trim(DataVencimento.ClipText)) = 0 Then gError 15799

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
        'Verifica se a DataContabil está preenchida
        If Len(Trim(DataContabil.ClipText)) = 0 Then gError 15798
    
        'Verifica se a DataContabil é maior ou igual que a DataEmissao
        If CDate(DataContabil.Text) < CDate(DataEmissao.Text) Then gError 15800

    End If
    
    'mover os dados da tela p/gobjBorderoCobranca
    gobjBorderoCobrancaEmissao.iCobrador = CInt(Codigo_Extrai(Cobrador.Text))
    gobjBorderoCobrancaEmissao.iCarteira = CInt(Codigo_Extrai(Carteira.Text))
    gobjBorderoCobrancaEmissao.dtDataEmissao = CDate(DataEmissao.Text)
    gobjBorderoCobrancaEmissao.dtDataVencimento = CDate(DataVencimento.Text)
    gobjBorderoCobrancaEmissao.dtDataContabil = MaskedParaDate(DataContabil)
    
    If Len(Trim(Totalizando.Text)) <> 0 Then

        gobjBorderoCobrancaEmissao.dValorMaximo = CDbl(Totalizando.Text)

    Else

        gobjBorderoCobrancaEmissao.dValorMaximo = 0

    End If
    
    If Len(Trim(Estado.Text)) <> 0 Then

        gobjBorderoCobrancaEmissao.sSiglaEstado = Estado.Text

    Else

        gobjBorderoCobrancaEmissao.sSiglaEstado = ""

    End If
    
    If Len(Trim(TipoCliente.Text)) <> 0 Then

        gobjBorderoCobrancaEmissao.iTipoCliente = Codigo_Extrai(TipoCliente.Text)

    Else
        gobjBorderoCobrancaEmissao.iTipoCliente = 0

    End If

    'garante que a colecao esteja vazia
    Do While gobjBorderoCobrancaEmissao.colInfoParcRec.Count > 0
        gobjBorderoCobrancaEmissao.colInfoParcRec.Remove (1)
    Loop
    
    '##############################################################
    'Inserido por Wagner
    lErro = CF2(Me, "BC_Move_Bordero_Memoria", gobjBorderoCobrancaEmissao)
    If lErro <> SUCESSO Then gError 140290
    '##############################################################
    
    lErro = CF("ParcelasRec_Le_BorderoCobranca", gobjBorderoCobrancaEmissao)
    If lErro <> SUCESSO Then gError 15801

    'Chama a tela do passo seguinte
    Call Chama_Tela("BorderoCobranca2", gobjBorderoCobrancaEmissao)
        
    'Fecha a tela
    Unload Me
    
    Exit Sub

Erro_BotaoSeguir_Click:

    Select Case gErr

        Case 15795
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr)
        
        Case 15796
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRA_COBRADOR_NAO_INFORMADA", gErr)
            
        Case 15797, 15798, 15799
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)
            
        Case 15800
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATACONTABIL_MENOR_DATAEMISSAO", gErr)

        Case 15801, 140290 'Inserido por Wagner
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155485)

    End Select

    Exit Sub

End Sub

Public Sub Carteira_Click()

Dim objCobrador As New ClassCobrador
Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Carteira_Click

    'Se a Carteira foi preenchida
    If Len(Trim(Carteira.Text)) <> 0 Then

        'Se é a Carteira selecionada na Combo
        If Carteira.Text = Carteira.List(Carteira.ListIndex) Then Exit Sub

        'Verifica se a Carteira existe na Combo. Se existir, seleciona
        lErro = Combo_Seleciona(Carteira, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 Then Error 15831

        'Se a Carteira(CODIGO) não existe na Combo
        If lErro = 6730 Then

            'Verifica se o Cobrador foi digitado
            If Len(Trim(Cobrador.Text)) = 0 Then Error 15832

            'Passa os Códigos da Carteira para o Obj
            objCarteiraCobranca.iCodigo = CInt(Carteira.Text)
            
            'Pesquisa se existe Carteira com o código em questão
            lErro = CF("CarteiraDeCobranca_Le", objCarteiraCobranca)
            If lErro <> SUCESSO And lErro <> 23413 Then Error 15833
            
            If lErro = 23413 Then Error 40628
            
        End If

    End If

    Exit Sub

Erro_Carteira_Click:

    Select Case Err

        Case 15830, 15831, 15833
            Carteira.SetFocus

        Case 15832
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", Err)
            Cobrador.SetFocus
            
        Case 40628
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA", Err, objCarteiraCobranca.iCodigo)
            Carteira.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155486)

    End Select

    Exit Sub

End Sub

Public Sub Cobrador_Click()

Dim iCodCobrador As Integer
Dim objCobrador As New ClassCobrador
Dim lErro As Long
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim sListBoxItem As String
Dim colCarteirasCobrador As New Collection

On Error GoTo Erro_Cobrador_Click
    
    'Limpa a Combo de Carteiras
    Carteira.Clear

    'Se Cobrador está preenchido
    If Len(Trim(Cobrador.Text)) <> 0 Then

        'Extrai o código do Cobrador
        iCodCobrador = Codigo_Extrai(Cobrador.Text)
    
        'Passa o Código do Cobrador que está na tela para o Obj
        objCobrador.iCodigo = iCodCobrador
    
        'Lê os dados do Cobrador
        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then gError 15824
    
        'Se o Cobrador não estiver cadastrado
        If lErro = 19294 Then gError 15825
                                
        'Le as carteiras associadas ao Cobrador
        lErro = CF("Cobrador_Le_Carteiras", objCobrador, colCarteirasCobrador)
        If lErro <> SUCESSO And lErro <> 23500 Then gError 40625

        If lErro = SUCESSO Then
        
            'Preencher a Combo
            For Each objCarteiraCobrador In colCarteirasCobrador
            
                If objCarteiraCobrador.iDesativada = DESMARCADO Then
                       
                    objCarteiraCobranca.iCodigo = objCarteiraCobrador.iCodCarteiraCobranca
         
                    lErro = CF("CarteiraDeCobranca_Le", objCarteiraCobranca)
                    If lErro <> SUCESSO And lErro <> 23413 Then gError 40626
         
                    'Carteira não está cadastrado
                    If lErro = 23413 Then gError 40627
        
                    'Concatena Código e a Descricao da carteira
                    sListBoxItem = CStr(objCarteiraCobranca.iCodigo)
                    sListBoxItem = sListBoxItem & SEPARADOR & objCarteiraCobranca.sDescricao
               
                    Carteira.AddItem sListBoxItem
                    Carteira.ItemData(Carteira.NewIndex) = objCarteiraCobranca.iCodigo
                
                End If
                
            Next
        End If
                
        'Seleciona uma das Carteiras
        If Carteira.ListCount <> 0 Then Carteira.Text = Carteira.List(0)
        If Carteira.ListCount = 0 Then gError 189566

    End If

    Exit Sub

Erro_Cobrador_Click:

    Select Case gErr

        Case 15824, 15836, 40625, 40626
            Cobrador.SetFocus
        
        Case 15825
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", gErr, Cobrador.Text)
            Cobrador.SetFocus
              
        Case 40627
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA", gErr, objCarteiraCobranca.iCodigo)
        
        Case 189566
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_SEM_CARTEIRA_ATIVA", gErr, objCobrador.iCodigo, objCobrador.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155487)

    End Select

    Exit Sub
    
End Sub

Function Trata_Parametros(Optional objBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se a tela é chamada do Menu então seta o objBordero e o preenche com os dados da tela
    If (objBorderoCobrancaEmissao Is Nothing) Then
        Set gobjBorderoCobrancaEmissao = New ClassBorderoCobrancaEmissao
            
        'Preenche os Obj(s) com os dados da tela
        gobjBorderoCobrancaEmissao.dtDataEmissao = CDate(DataEmissao.Text)
        gobjBorderoCobrancaEmissao.dtDataContabil = StrParaDate(DataContabil.Text)
        gobjBorderoCobrancaEmissao.dtDataVencimento = CDate(DataVencimento.Text)
        
        If Cobrador.ListCount <> 0 Then
            gobjBorderoCobrancaEmissao.iCobrador = Cobrador.ItemData(Cobrador.ListIndex)
            gobjBorderoCobrancaEmissao.iCarteira = Carteira.ItemData(Cobrador.ListIndex)
        End If
        
        gobjBorderoCobrancaEmissao.iTipoCliente = Codigo_Extrai(TipoCliente.Text)
        gobjBorderoCobrancaEmissao.sSiglaEstado = Estado.Text
        gobjBorderoCobrancaEmissao.dValorMaximo = StrParaDbl(Totalizando.Text)
               
    'Se tela for chamada de outro ponto então preenche os campos com os
    'Com os dados do Obj
    Else
        Set gobjBorderoCobrancaEmissao = objBorderoCobrancaEmissao
        
        'Preenche a ComboBox Cobrador
        For iIndice = 0 To Cobrador.ListCount - 1
            If Cobrador.ItemData(iIndice) = objBorderoCobrancaEmissao.iCobrador Then
                Cobrador.ListIndex = iIndice
                Exit For
            End If
        Next
    
        'Preenche a ComboBox Carteira
        For iIndice = 0 To Carteira.ListCount - 1
            If Carteira.ItemData(iIndice) = objBorderoCobrancaEmissao.iCarteira Then
                Carteira.ListIndex = iIndice
                Exit For
            End If
        Next
        
        'Preenche a ComboBox TipoCliente
        For iIndice = 0 To TipoCliente.ListCount - 1
            If TipoCliente.ItemData(iIndice) = objBorderoCobrancaEmissao.iTipoCliente Then
                TipoCliente.ListIndex = iIndice
                Exit For
            End If
        Next
        
        'Preenche os demais campos da tela
        DataEmissao.Text = Format(objBorderoCobrancaEmissao.dtDataEmissao, "dd/mm/yy")
        Call DateParaMasked(DataContabil, objBorderoCobrancaEmissao.dtDataContabil)
        DataVencimento.Text = Format(objBorderoCobrancaEmissao.dtDataVencimento, "dd/mm/yy")
        Estado.Text = objBorderoCobrancaEmissao.sSiglaEstado
                
        '##############################################################
        'Inserido por Wagner
        lErro = CF2(Me, "BC_Traz_Bordero_Tela", gobjBorderoCobrancaEmissao)
        If lErro <> SUCESSO Then gError 140291
        '##############################################################
                
        'Verifica se dValorMaximo não é vazio
        If objBorderoCobrancaEmissao.dValorMaximo <> 0 Then
            Totalizando.Text = Format(objBorderoCobrancaEmissao.dValorMaximo, "standard")
        End If
    
    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 15814
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARAMETRO_OBRIGATORIO", gErr)
            
        Case 140291

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155488)

    End Select

    Exit Function

End Function

Public Sub DataContabil_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataContabil)

End Sub

Public Sub DataContabil_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataContabil_Validate

    'Verifica se a data está preenchida
    If Len(Trim(DataContabil.ClipText)) = 0 Then Exit Sub

    'Verifica se a data é válida
    lErro = Data_Critica(DataContabil.Text)
    If lErro <> SUCESSO Then Error 41624

    Exit Sub

Erro_DataContabil_Validate:

    Cancel = True


    Select Case Err

        Case 41624

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155489)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao)

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a data de emissão está preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Verifica se a data é válida
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then Error 41622

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case Err

        Case 41622

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155490)

    End Select

    Exit Sub

End Sub

Public Sub DataVencimento_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataVencimento)

End Sub

Public Sub DataVencimento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataVencimento_Validate

    'Verifica se a data está preenchida
    If Len(Trim(DataVencimento.ClipText)) = 0 Then Exit Sub

    'Verifica se a data é válida
    lErro = Data_Critica(DataVencimento.Text)
    If lErro <> SUCESSO Then Error 41623

    Exit Sub

Erro_DataVencimento_Validate:

    Cancel = True


    Select Case Err

        Case 41623

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155491)

    End Select

    Exit Sub

End Sub

Public Sub Estado_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Estado_Validate

    'Verifica se foi preenchido o Estado
    If Len(Trim(Estado.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Estado
    If Estado.Text = Estado.List(Estado.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na Combo Estado, se existir seleciona o item
    lErro = Combo_Item_Igual_CI(Estado)
    If lErro <> SUCESSO And lErro <> 58583 Then Error 40623

    'Não existe o ítem na ComboBox Estado
    If lErro = 58583 Then Error 40624

    Exit Sub

Erro_Estado_Validate:

    Cancel = True


    Select Case Err

    Case 40623

    Case 40624
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", Err, Estado.Text)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155492)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long, iIndice As Integer
Dim sCodigoCobrador As String, sCodigoEstado As String
Dim sCodigoCliente As String, sNomeCobrador As String
Dim sNomeEstado As String, sNomeCliente As String
Dim vCodigoEstado As Variant, colCodigo As New Collection
Dim colCodigoDescricaoCobradores As New AdmColCodigoNome
Dim colCodigoDescricaoEstados As New AdmColCodigoNome
Dim colCodigoDescricaoClientes As New AdmColCodigoNome
Dim objCodigoNomeCobradores As AdmCodigoNome
Dim objCodigoNomeEstados As AdmCodigoNome
Dim objCodigoNomeClientes As AdmCodigoNome
Dim ColCobrador As New Collection
Dim objCobrador As ClassCobrador

On Error GoTo Erro_Form_Load
    
    sCodigoEstado = "Sigla"
    sNomeEstado = "Nome"
    sCodigoCobrador = "Codigo"
    sNomeCobrador = "NomeReduzido"
    sCodigoCliente = "Codigo"
    sNomeCliente = "Descricao"

    'Carrega a Coleção de Cobradores
    
    lErro = CF("Cobradores_Le_Todos_Filial", ColCobrador)
    If lErro <> SUCESSO Then Error 15815
    
    For Each objCobrador In ColCobrador
    
        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA And objCobrador.iInativo <> Inativo Then
        
            Cobrador.AddItem objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
            Cobrador.ItemData(Cobrador.NewIndex) = objCobrador.iCodigo
            
        End If

    Next
     
    'Se houver cobradores cadastrados
    If ColCobrador.Count <> 0 Then
        'Seleciona um dos Cobradores
        If Cobrador.ListCount <> 0 Then Cobrador.Text = Cobrador.List(0)
    End If
    
    'Lê cada código da tabela Estados
    lErro = CF("Codigos_Le", "Estados", sCodigoEstado, TIPO_STR, colCodigo, STRING_ESTADOS_SIGLA)
    If lErro <> SUCESSO Then Error 15835

    'Preenche as ComboBox Estados com os objetos da colecao colCodigo
    For Each vCodigoEstado In colCodigo
        
        Estado.AddItem vCodigoEstado
    
    Next
    
    'Carrega a Coleção de Tipos de Cliente
    lErro = CF("Cod_Nomes_Le", "TiposDeCliente", sCodigoCliente, sNomeCliente, STRING_TIPO_CLIENTE_DESCRICAO, colCodigoDescricaoClientes)
    If lErro <> SUCESSO Then Error 15817
    
    'Preenche a ComboBox Tipo Cliente com os objetos da coleção de Tipos de Cliente
    For Each objCodigoNomeClientes In colCodigoDescricaoClientes

        TipoCliente.AddItem CStr(objCodigoNomeClientes.iCodigo) & SEPARADOR & objCodigoNomeClientes.sNome
        TipoCliente.ItemData(TipoCliente.NewIndex) = objCodigoNomeClientes.iCodigo

    Next

    'Preenche as Datas com a data corrente do sistema
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataVencimento.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    If Not (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
        
        DataContabil.Enabled = False
        LabelDataContabil.Enabled = False
        UpDownDataContabil.Enabled = False
        
    Else
    
        DataContabil.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    End If
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 15815, 15816, 15817, 15835

        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155493)

    End Select
    
    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set gobjBorderoCobrancaEmissao = Nothing
    
End Sub

Public Sub TipoCliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoCliente As New ClassTipoCliente
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_TipoCliente_Validate

    'Verifica se foi preenchida a ComboBox Tipo
    If Len(Trim(TipoCliente.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Tipo
    If TipoCliente.Text = TipoCliente.List(TipoCliente.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TipoCliente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19219

    'Não existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTipoCliente.iCodigo = iCodigo

        'Tenta ler TipoCliente com esse código no BD
        lErro = CF("TipoDeCliente_Le", objTipoCliente)
        If lErro <> SUCESSO And lErro <> 28943 Then Error 19220

        'Não encontrou Tipo Cliente no BD
        If lErro = 28943 Then Error 33000

        'Encontrou TipoCliente no BD, coloca no Text da Combo
        TipoCliente.Text = CStr(iCodigo) & SEPARADOR & objTipoCliente.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19222

    Exit Sub

Erro_TipoCliente_Validate:

    Cancel = True


    Select Case Err

        Case 19219, 19220  'Já tratado na rotina chamada


        Case 19222, 33000 'Não encontrou Tipo Cliente no BD

            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCLIENTE_INEXISTENTE", Err)

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155494)

    End Select

    Exit Sub

End Sub

Public Sub Totalizando_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Totalizando_Validate

    'Se Totalizando está preenchido
    If Len(Trim(Totalizando.Text)) > 0 Then

        'Verifica se Totalizando é válido
        lErro = Valor_Positivo_Critica(Totalizando.Text)
        If lErro <> SUCESSO Then Error 15851

        'Formata o texto
        Totalizando.Text = Format(Totalizando.Text, "Standard")

    End If

    Exit Sub

Erro_Totalizando_Validate:

    Cancel = True


    Select Case Err

        Case 15851

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155495)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataContabil_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataContabil_DownClick

    'Diminui a DataContabil em 1 dia
    lErro = Data_Up_Down_Click(DataContabil, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15823

    Exit Sub

Erro_UpDownDataContabil_DownClick:

    Select Case Err

        Case 15823

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155496)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataContabil_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataContabil_UpClick

    'Aumenta a DataContabil em 1 dia
    lErro = Data_Up_Down_Click(DataContabil, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15822

    Exit Sub

Erro_UpDownDataContabil_UpClick:

    Select Case Err

        Case 15822

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155497)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataEmissao_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_DownClick

    'Diminui a DataEmissao em 1 dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15821

    Exit Sub

Erro_UpDownDataEmissao_DownClick:

    Select Case Err

        Case 15821

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155498)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataEmissao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_UpClick

    'Aumenta a DataEmissao em 1 dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15820

    Exit Sub

Erro_UpDownDataEmissao_UpClick:

    Select Case Err

        Case 15820

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155499)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataVencimento_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimento_DownClick

    'Diminui a DataVencimento em 1 dia
    lErro = Data_Up_Down_Click(DataVencimento, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15819

    Exit Sub

Erro_UpDownDataVencimento_DownClick:

    Select Case Err

        Case 15819

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155500)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataVencimento_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimento_UpClick

    'Aumenta a DataVencimento em 1 dia
    lErro = Data_Up_Down_Click(DataVencimento, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15818

    Exit Sub

Erro_UpDownDataVencimento_UpClick:

    Select Case Err

        Case 15818

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155501)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BORDERO_COBRANCA
    Set Form_Load_Ocx = Me
    Caption = "Bordero de Cobrança"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BorderoCobranca"
    
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

'######################################################################
'Inserido por Wagner

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then


    End If

End Sub

Function BC_Move_Bordero_Memoria(ByVal objCT As Object, ByVal objBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao) As Long
    BC_Move_Bordero_Memoria = SUCESSO
End Function
Function BC_Traz_Bordero_Tela(ByVal objCT As Object, ByVal objBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao) As Long
    BC_Traz_Bordero_Tela = SUCESSO
End Function
'######################################################################

