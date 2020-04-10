VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTabEnderecoBol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objTela As Object
Private colTab As Collection
Private iTabAtual As Integer
Private iIndexBrasil As Integer
Private iIndexUF As Integer
Private bTelaFilialEmp As Boolean
Private sCEPAnt(1 To 5) As String
Private bMudouCEP As Boolean

Private WithEvents objEventoPais As AdmEvento
Attribute objEventoPais.VB_VarHelpID = -1
Private WithEvents objEventoCidade As AdmEvento
Attribute objEventoCidade.VB_VarHelpID = -1

Public Function Finaliza()
    Set objEventoPais = Nothing
    Set objEventoCidade = Nothing
    Set colTab = Nothing
    Set objTela = Nothing
End Function

Public Function Inicializa(objTelaAux As Object, objTab1 As Object, Optional objTab2 As Object = Nothing, Optional objTab3 As Object = Nothing, Optional objTab4 As Object = Nothing, Optional objTab5 As Object = Nothing) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigo As New Collection
Dim vCodigo As Variant
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objTab As Object, objControl As Object

On Error GoTo Erro_Inicializa
    
    Set objTela = objTelaAux
    Set colTab = New Collection
    Set objEventoPais = New AdmEvento
    Set objEventoCidade = New AdmEvento
    
    bMudouCEP = False
    
    If InStr(1, objTelaAux.Name, "FilialEmpresa") <> 0 Then
        bTelaFilialEmp = True
    Else
        bTelaFilialEmp = False
    End If
    
    Set objTab1.gobjTela = objTela
    objTab1.giindex = 1
    colTab.Add objTab1
    If Not (objTab2 Is Nothing) Then
        Set objTab2.gobjTela = objTela
        objTab2.giindex = 2
        colTab.Add objTab2
    End If
    If Not (objTab3 Is Nothing) Then
        Set objTab3.gobjTela = objTela
        objTab3.giindex = 3
        colTab.Add objTab3
    End If
    If Not (objTab4 Is Nothing) Then
        Set objTab4.gobjTela = objTela
        objTab4.giindex = 4
        colTab.Add objTab4
    End If
    If Not (objTab5 Is Nothing) Then
        Set objTab5.gobjTela = objTela
        objTab5.giindex = 5
        colTab.Add objTab5
    End If
       
    'Lê cada codigo da tabela Estados
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colCodigo, STRING_ESTADOS_SIGLA)
    If lErro <> SUCESSO Then gError 202958
    
    'Lê cada codigo e descricao da tabela Paises
    lErro = CF("Cod_Nomes_Le", "Paises", "Codigo", "Nome", STRING_PAISES_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 202959
    
    For Each objTab In colTab
    
        If bTelaFilialEmp Then
            Controle("PaisLabel", objTab.giindex).MousePointer = vbDefault
            Controle("LabelCidade", objTab.giindex).MousePointer = vbDefault
        Else
            Controle("PaisLabel", objTab.giindex).MousePointer = vbArrowQuestion
            Controle("LabelCidade", objTab.giindex).MousePointer = vbArrowQuestion
        End If
    
        Controle("TipoLogradouro", objTab.giindex).Clear
    
        lErro = CF("Carrega_Combo", Controle("TipoLogradouro", objTab.giindex), "TiposDeLogradouro", "Sigla", TIPO_STR, "Nome", TIPO_STR)
        If lErro <> SUCESSO Then gError 202959
    
        Controle("Estado", objTab.giindex).Clear
        
        'Preenche as ComboBox Estados com os objetos da colecao colCodigo
        For Each vCodigo In colCodigo
            Controle("Estado", objTab.giindex).AddItem vCodigo
        Next
        
        'Preenche cada ComboBox País com os objetos da colecao colCodigoDescricao
        For Each objCodigoDescricao In colCodigoDescricao
            Controle("Pais", objTab.giindex).AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
            Controle("Pais", objTab.giindex).ItemData(Controle("Pais", objTab.giindex).NewIndex) = objCodigoDescricao.iCodigo
        Next
        
        Controle("Logradouro", objTab.giindex).MaxLength = STRING_ENDERECO
        Controle("Bairro", objTab.giindex).MaxLength = STRING_BAIRRO
        Controle("Cidade", objTab.giindex).MaxLength = STRING_CIDADE
        Controle("TelNumero1", objTab.giindex).MaxLength = STRING_TELEFONE
        Controle("TelNumero2", objTab.giindex).MaxLength = STRING_TELEFONE
        Controle("FaxNumero", objTab.giindex).MaxLength = STRING_FAX
        Controle("Email1", objTab.giindex).MaxLength = STRING_EMAIL
        Controle("Email2", objTab.giindex).MaxLength = STRING_EMAIL
        Controle("Contato", objTab.giindex).MaxLength = STRING_CONTATO
        Controle("Skype", objTab.giindex).MaxLength = STRING_SKYPE
        Controle("Radio", objTab.giindex).MaxLength = STRING_RADIO
        
    Next
    
    'Seleciona Brasil se existir
    For iIndice = 0 To Controle("Pais").ListCount - 1
        If right(Controle("Pais").List(iIndice), 6) = "Brasil" Then
            iIndexBrasil = iIndice
            Exit For
        End If
    Next
    
    Call Limpa_Tela
    
    Inicializa = SUCESSO

    Exit Function

Erro_Inicializa:

    Inicializa = gErr
    
    Select Case gErr
    
        Case 202958, 202959
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202957)

    End Select

    Exit Function
    
End Function

Public Function Move_Endereco_Memoria(ByVal colEndereco As Collection) As Long

Dim lErro As Long
Dim objEndereco As ClassEndereco
Dim objEnderecoBol As ClassEnderecoBol
Dim objTab As Object

On Error GoTo Erro_Move_Endereco_Memoria

    For Each objTab In colTab

        Set objEndereco = New ClassEndereco
        Set objEnderecoBol = New ClassEnderecoBol

        objEndereco.sLogradouro = Trim(Controle("Logradouro", objTab.giindex).Text)
        objEndereco.sTipoLogradouro = SCodigo_Extrai(Controle("TipoLogradouro", objTab.giindex).Text)
        objEndereco.sComplemento = Trim(Controle("Complemento", objTab.giindex).Text)
        objEndereco.lNumero = StrParaLong(Controle("Numero", objTab.giindex).Text)
        objEndereco.sBairro = Trim(Controle("Bairro", objTab.giindex).Text)
        objEndereco.sCidade = Trim(Controle("Cidade", objTab.giindex).Text)
        objEndereco.sCEP = Trim(Controle("CEP", objTab.giindex).Text)
        objEndereco.sReferencia = Trim(Controle("Referencia", objTab.giindex).Text)

        objEndereco.iCodigoPais = Codigo_Extrai(Controle("Pais", objTab.giindex).Text)
        objEndereco.sSiglaEstado = Trim(Controle("Estado", objTab.giindex).Text)
        If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = PAIS_BRASIL
        If objEndereco.iCodigoPais = PAIS_BRASIL And (Controle("Estado", objTab.giindex).ListIndex = -1 Or Len(Trim(Controle("Estado", objTab.giindex).Text)) = 0) Then gError 12325

        objEndereco.sTelNumero1 = Trim(Controle("TelNumero1", objTab.giindex).Text)
        objEndereco.sTelNumero2 = Trim(Controle("TelNumero2", objTab.giindex).Text)
        objEndereco.sFaxNumero = Trim(Controle("FaxNumero", objTab.giindex).Text)
        objEndereco.iTelDDD1 = StrParaInt(Controle("TelDDD1", objTab.giindex).Text)
        objEndereco.iTelDDD2 = StrParaInt(Controle("TelDDD2", objTab.giindex).Text)
        objEndereco.iFaxDDD = StrParaInt(Controle("FaxDDD", objTab.giindex).Text)
        
        objEndereco.sEmail = Trim(Controle("Email1", objTab.giindex).Text)
        objEndereco.sEmail2 = Trim(Controle("Email2", objTab.giindex).Text)
        objEndereco.sContato = Trim(Controle("Contato", objTab.giindex).Text)
        
        objEnderecoBol.sRadio = Trim(Controle("Radio", objTab.giindex).Text)
        objEnderecoBol.sSkype = Trim(Controle("Skype", objTab.giindex).Text)
        
        Set objEndereco.objInfoUsu = objEnderecoBol

        colEndereco.Add objEndereco

    Next
    
    Move_Endereco_Memoria = SUCESSO

    Exit Function

Erro_Move_Endereco_Memoria:

    Move_Endereco_Memoria = gErr

    Select Case gErr
    
        Case 12325
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, objEndereco.sSiglaEstado)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202956)

    End Select

    Exit Function

End Function

Public Function Traz_Endereco_Tela(ByVal colEndereco As Object) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objEndereco As ClassEndereco
Dim objTab As Object
Dim objEnderecoBol As ClassEnderecoBol

On Error GoTo Erro_Traz_Endereco_Tela

    If colEndereco.Count <> colTab.Count Then gError 202954
    
    iIndice = 0
    For Each objEndereco In colEndereco
       
        iIndice = iIndice + 1
        Set objTab = colTab.Item(iIndice)
        
        lErro = CF("Endereco_Le_Customizado", objEndereco)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If Not (objEndereco.objInfoUsu Is Nothing) Then
            Set objEnderecoBol = objEndereco.objInfoUsu
            Controle("Radio", objTab.giindex).Text = objEnderecoBol.sRadio
            Controle("Skype", objTab.giindex).Text = objEnderecoBol.sSkype
        Else
            Controle("Radio", objTab.giindex).Text = ""
            Controle("Skype", objTab.giindex).Text = ""
        End If

        Controle("Pais", objTab.giindex).Text = objEndereco.iCodigoPais
        Call Pais_Validate(bSGECancelDummy, objTab.giindex)
        Controle("Bairro", objTab.giindex).Text = objEndereco.sBairro
        Controle("Cidade", objTab.giindex).Text = objEndereco.sCidade
        sCEPAnt(iIndice) = objEndereco.sCEP
        Controle("CEP", objTab.giindex).Text = objEndereco.sCEP
        Controle("Estado", objTab.giindex).Text = objEndereco.sSiglaEstado
        Call Estado_Validate(bSGECancelDummy, objTab.giindex)
        If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = PAIS_BRASIL
    
        Controle("Email1", objTab.giindex).Text = objEndereco.sEmail
        Controle("Email2", objTab.giindex).Text = objEndereco.sEmail2
        Controle("Contato", objTab.giindex).Text = objEndereco.sContato
    
        Controle("TipoLogradouro", objTab.giindex).Text = objEndereco.sTipoLogradouro
        Call TipoLogradouro_Validate(bSGECancelDummy, objTab.giindex)
        
        Controle("Logradouro", objTab.giindex).Text = objEndereco.sLogradouro
        
        If objEndereco.lNumero <> 0 Then
            Controle("Numero", objTab.giindex).Text = CStr(objEndereco.lNumero)
        Else
            Controle("Numero", objTab.giindex).Text = ""
        End If
        Controle("Complemento", objTab.giindex).Text = objEndereco.sComplemento
        If objEndereco.iTelDDD1 <> 0 Then
            Controle("TelDDD1", objTab.giindex).Text = Format(objEndereco.iTelDDD1, "00")
        Else
            Controle("TelDDD1", objTab.giindex).Text = ""
        End If
        If objEndereco.iTelDDD2 <> 0 Then
            Controle("TelDDD2", objTab.giindex).Text = Format(objEndereco.iTelDDD2, "00")
        Else
            Controle("TelDDD2", objTab.giindex).Text = ""
        End If
        Controle("TelNumero1", objTab.giindex).Text = objEndereco.sTelNumero1
        Controle("TelNumero2", objTab.giindex).Text = objEndereco.sTelNumero2
        If objEndereco.iFaxDDD <> 0 Then
            Controle("FaxDDD", objTab.giindex).Text = Format(objEndereco.iFaxDDD, "00")
        Else
            Controle("FaxDDD", objTab.giindex).Text = ""
        End If
        Controle("FaxNumero", objTab.giindex).Text = objEndereco.sFaxNumero
        Controle("Referencia", objTab.giindex).Text = objEndereco.sReferencia
        
    Next

    Traz_Endereco_Tela = SUCESSO

    Exit Function

Erro_Traz_Endereco_Tela:

    Traz_Endereco_Tela = gErr

    Select Case gErr
    
        Case 202954
            Call Rotina_Erro(vbOKOnly, "ERRO_END_TAB_DIF", gErr, colEndereco.Count, colTab.Count)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202955)

    End Select

    Exit Function

End Function

Public Function Limpa_Tela()

Dim objTab As Object
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela
    
    'Seleciona Brasil nas Combos País
    For Each objTab In colTab
        Call objTab.object.Limpa_Tab
        Controle("Pais", objTab.giindex).ListIndex = iIndexBrasil
        Controle("TipoLogradouro", objTab.giindex).ListIndex = -1
    Next
    
    Call FilialEmpresa_SetaEstados
    For iIndice = LBound(sCEPAnt) To UBound(sCEPAnt)
        sCEPAnt(iIndice) = ""
    Next
    bMudouCEP = False
    
    Limpa_Tela = SUCESSO
    
    Exit Function

Erro_Limpa_Tela:

    Limpa_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202960)

    End Select

    Exit Function
End Function

Private Function FilialEmpresa_SetaEstados()

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objEndereco As New ClassEndereco
Dim iSigla As Integer
Dim iIndice As Integer
Dim objTab As Object

On Error GoTo Erro_FilialEmpresa_SetaEstados

    If giFilialEmpresa <> EMPRESA_TODA And iIndexUF = 0 Then

        objFilialEmpresa.iCodFilial = giFilialEmpresa
        'Lê a Filial Empresa
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 202961

        'Se não encontrou a Filial da Empresa --> Erro
        If lErro <> SUCESSO Then gError 202962

        Set objEndereco = objFilialEmpresa.objEndereco

        If objEndereco.sSiglaEstado <> "" Then

            iSigla = Len(Trim(objEndereco.sSiglaEstado))
            'Seleciona o Estado "default" p/ a Filial se existir
            For iIndice = 0 To Controle("Estado").ListCount - 1
                If UCase(right(Controle("Estado").List(iIndice), iSigla)) = UCase(objEndereco.sSiglaEstado) Then
                    iIndexUF = iIndice
                    Exit For
                End If
            Next

        End If

    End If
    
    For Each objTab In colTab
        Controle("Estado", objTab.giindex).ListIndex = iIndexUF
    Next

    FilialEmpresa_SetaEstados = SUCESSO

    Exit Function

Erro_FilialEmpresa_SetaEstados:

    FilialEmpresa_SetaEstados = gErr

    Select Case gErr

        Case 202961

        Case 202962
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202963)

    End Select

    Exit Function

End Function

Private Function Controle(ByVal sCampo As String, Optional ByVal iIndice As Integer = 0) As Object

On Error GoTo Erro_Controle

    If iIndice = 0 Then iIndice = 1
    
    iTabAtual = iIndice
    
    Set Controle = colTab.Item(iIndice).object.Controls(sCampo)
    
    Exit Function

Erro_Controle:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202964)

    End Select

    Exit Function
    
End Function

Public Sub Bairro_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CEP_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CEP_GotFocus(ByVal iIndice As Integer)
    Call MaskEdBox_TrataGotFocus(Controle("CEP", iIndice), objTela.iAlterado)
End Sub

Public Sub Cidade_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Contato_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Email1_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Email2_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Skype_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Radio_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Email1_Validate(Cancel As Boolean, ByVal iIndice As Integer)
'
End Sub

Public Sub Email2_Validate(Cancel As Boolean, ByVal iIndice As Integer)
'
End Sub

Public Sub Logradouro_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoLogradouro_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoLogradouro_Click(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Numero_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Numero_GotFocus(ByVal iIndice As Integer)
    Call MaskEdBox_TrataGotFocus(Controle("Numero", iIndice), objTela.iAlterado)
End Sub

Public Sub Complemento_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Referencia_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TelDDD1_GotFocus(ByVal iIndice As Integer)
    Call MaskEdBox_TrataGotFocus(Controle("TelDDD1", iIndice), objTela.iAlterado)
End Sub

Public Sub TelDDD2_GotFocus(ByVal iIndice As Integer)
    Call MaskEdBox_TrataGotFocus(Controle("TelDDD2", iIndice), objTela.iAlterado)
End Sub

Public Sub FaxDDD_GotFocus(ByVal iIndice As Integer)
    Call MaskEdBox_TrataGotFocus(Controle("FaxDDD", iIndice), objTela.iAlterado)
End Sub

Public Sub FaxDDD_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TelDDD1_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TelDDD2_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FaxNumero_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TelNumero1_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TelNumero2_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Estado_Click(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
    Call Trata_Alteracao_UF(iIndice)
End Sub

Public Sub Estado_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Estado_Validate(Cancel As Boolean, ByVal iIndice As Integer)

Dim lErro As Long

On Error GoTo Erro_Estado_Validate

    'Verifica se foi preenchido o Estado
    If Len(Trim(Controle("Estado", iIndice).Text)) <> 0 Then

        'Verifica se está preenchida com o ítem selecionado na ComboBox Estado
        If Controle("Estado", iIndice).Text = Controle("Estado", iIndice).List(Controle("Estado", iIndice).ListIndex) Then Exit Sub
    
        'Verifica se existe o ítem na Combo Estado, se existir seleciona o item
        lErro = Combo_Item_Igual_CI(Controle("Estado", iIndice))
        If lErro <> SUCESSO And lErro <> 58583 Then gError 202965
    
        'Não existe o ítem na ComboBox Estado
        If lErro = 58583 And Codigo_Extrai(Controle("Pais", iIndice)) = PAIS_BRASIL Then gError 202966
    
    End If
    
    Call Trata_Alteracao_UF(iIndice)

    Exit Sub

Erro_Estado_Validate:

    Cancel = True

    Select Case gErr

        Case 202965

        Case 202966
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, Controle("Estado", iIndice).Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202967)

    End Select

    Exit Sub

End Sub

Public Sub Trata_Alteracao_UF(ByVal iIndice As Integer)

Dim objTab As Object
Dim sLog As String
Dim iPaisIndex As Integer
Dim iEstadoIndex As Integer

On Error GoTo Erro_Trata_Alteracao_UF

    'Se for o endereço principal
    If iIndice = 1 Then

        sLog = Controle("Logradouro", iIndice).Text
        iPaisIndex = Controle("Pais", iIndice).ListIndex
        iEstadoIndex = Controle("Estado", iIndice).ListIndex
    
        'Se tiver o endereço preenchido
        'If Len(Trim(sLog)) > 0 Then
            For Each objTab In colTab
            
                'Para os outros endereços
                If iIndice <> objTab.giindex Then
                
                    sLog = Controle("Logradouro", objTab.giindex).Text
                    
                    'Se não tiverem o endereço preenchido, mexe na UF e Paisa de acordo com o tab inicial
                    If Len(Trim(sLog)) = 0 Then
                        Controle("Pais", objTab.giindex).ListIndex = iPaisIndex
                        Controle("Estado", objTab.giindex).ListIndex = iEstadoIndex
                    End If
                End If
            
            Next
        'End If
        
    End If
    
    For Each objTab In colTab
        If Codigo_Extrai(Controle("Pais", objTab.giindex)) = PAIS_BRASIL Then
            Controle("Estado", objTab.giindex).Enabled = True
            If Controle("Estado", objTab.giindex).Text = "EX" Then Controle("Estado", objTab.giindex).ListIndex = iIndexUF
        Else
            Controle("Estado", objTab.giindex).Enabled = False
            Controle("Estado", objTab.giindex).Text = "EX"
        End If
    Next

    Exit Sub

Erro_Trata_Alteracao_UF:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202967)

    End Select

    Exit Sub

End Sub

Public Sub TipoLogradouro_Validate(Cancel As Boolean, ByVal iIndice As Integer, Optional ByVal bSemErro As Boolean = False)

Dim lErro As Long
Dim iIndiceAux As Integer
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim iIndice3 As Integer
Dim iIndice4 As Integer
Dim iTam As Integer
Dim iPos As Integer
Dim sValor As String
Dim sValorRed As String
Dim sValorLin As String

On Error GoTo Erro_TipoLogradouro_Validate

    'Verifica se foi preenchido o TipoLogradouro
    If Len(Trim(Controle("TipoLogradouro", iIndice).Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox TipoLogradouro
    If UCase(SCodigo_Extrai(Controle("TipoLogradouro", iIndice).Text)) = UCase(SCodigo_Extrai(Controle("TipoLogradouro", iIndice).List(Controle("TipoLogradouro", iIndice).ListIndex))) Then Exit Sub

    'Verifica se existe o ítem na Combo TipoLogradouro, se existir seleciona o item
    lErro = Combo_Item_Igual_CI(Controle("TipoLogradouro", iIndice))
    If lErro <> SUCESSO And lErro <> 58583 Then gError 202978
    
    sValor = UCase(SCodigo_Extrai(Controle("TipoLogradouro", iIndice).Text))
    iTam = Len(sValor)
    iPos = InStr(1, sValor, " ")
    If iPos <> 0 Then
        sValorRed = left(sValor, iPos - 1)
    Else
        sValorRed = ""
    End If

    'Não existe o ítem na ComboBox TipoLogradouro
    If lErro = 58583 Then 'gError 202966
    
        iIndice1 = -1
        iIndice2 = -1
        iIndice3 = -1
        iIndice4 = -1
        For iIndiceAux = 0 To Controle("TipoLogradouro", iIndice).ListCount - 1
        
            sValorLin = Controle("TipoLogradouro", iIndice).List(iIndiceAux)
        
            If sValor = UCase(SCodigo_Extrai(sValorLin)) Then
                    iIndice1 = iIndiceAux
                Exit For
            End If
        
            If sValor = left(UCase(SCodigo_Extrai(sValorLin)), iTam) Then
                    If iIndice2 = -1 Then iIndice2 = iIndiceAux
            End If
            
            iPos = InStr(1, sValorLin, SEPARADOR)
            
            If sValor = Mid(UCase(sValorLin), iPos + 1, iTam) Then
                    If iIndice3 = -1 Then iIndice3 = iIndiceAux
            End If

            If sValorRed = Mid(UCase(sValorLin), iPos + 1, iTam) Then
                    If iIndice4 = -1 Then iIndice4 = iIndiceAux
            End If
            
        Next
        
        If iIndice1 = -1 And iIndice2 = -1 And iIndice3 = -1 And iIndice4 = -1 Then
            If Not bSemErro Then gError 202979
            Controle("TipoLogradouro", iIndice) = "OUT-OUTROS"
        End If
    
    
        If iIndice1 <> -1 Then
            Controle("TipoLogradouro", iIndice).ListIndex = iIndice1
        ElseIf iIndice2 <> -1 Then
            Controle("TipoLogradouro", iIndice).ListIndex = iIndice2
        ElseIf iIndice3 <> -1 Then
            Controle("TipoLogradouro", iIndice).ListIndex = iIndice3
        Else
            Controle("TipoLogradouro", iIndice).ListIndex = iIndice4
        End If
    
    End If
    
    Exit Sub

Erro_TipoLogradouro_Validate:

    Cancel = True

    Select Case gErr

        Case 202978

        Case 202979
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOLOGRADOURO_NAO_CADASTRADO", gErr, Controle("TipoLogradouro", iIndice).Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202980)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPais_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPais As New ClassPais

On Error GoTo Erro_objEventoPais_evSelecao

    Set objPais = obj1

    Controle("Pais", iTabAtual).Text = CStr(objPais.iCodigo)
    Call Pais_Validate(bSGECancelDummy, iTabAtual)

    objTela.Show

    Exit Sub

Erro_objEventoPais_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202968)

    End Select

    Exit Sub

End Sub

Public Sub Pais_Change(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Pais_Click(ByVal iIndice As Integer)
    objTela.iAlterado = REGISTRO_ALTERADO
    Call Trata_Alteracao_UF(iIndice)
End Sub

Public Sub Pais_Validate(Cancel As Boolean, ByVal iIndice As Integer)

Dim lErro As Long
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objPais As New ClassPais

On Error GoTo Erro_Pais_Validate

    'Verifica se foi preenchida a Combo Pais
    If Len(Trim(Controle("Pais", iIndice).Text)) <> 0 Then

    '    'Verifica se está preenchida com o ítem selecionado na ComboBox Pais
    '    If Controle("Pais", iIndice).Text = Controle("Pais", iIndice).List(Controle("Pais", iIndice).ListIndex) Then Exit Sub
    
        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Controle("Pais", iIndice), iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 202969
    
        'Nao existe o item com o CODIGO na List da ComboBox
        If lErro = 6730 Then
    
            objPais.iCodigo = iCodigo
    
            'Tenta ler Pais com esse codigo no BD
            lErro = CF("Paises_Le", objPais)
            If lErro <> SUCESSO And lErro <> 47876 Then gError 202970
            If lErro <> SUCESSO Then gError 202971
    
            Controle("Pais", iIndice).Text = CStr(iCodigo) & SEPARADOR & objPais.sNome
    
        End If
    
        'Nao existe o item com a STRING na List da ComboBox
        If lErro = 6731 Then gError 202972
        
        If Codigo_Extrai(Controle("Pais", iIndice)) = PAIS_BRASIL Then
            Controle("Estado", iIndice).Enabled = True
            If Controle("Estado", iIndice).Text = "EX" Then Controle("Estado", iIndice).ListIndex = iIndexUF
        Else
            Controle("Estado", iIndice).Enabled = False
            Controle("Estado", iIndice).Text = "EX"
        End If
        
    End If
    
    Call Trata_Alteracao_UF(iIndice)

    Exit Sub

Erro_Pais_Validate:

    Cancel = True

    Select Case gErr
    
        Case 202969, 202970

        Case 202971
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PAIS", objPais.iCodigo)
            If vbMsgRes = vbYes Then Call Chama_Tela("Paises", objPais)

        Case 202972
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO1", gErr, Trim(Controle("Pais", iIndice).Text))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202973)

    End Select

    Exit Sub

End Sub

Public Sub PaisLabel_Click(ByVal iIndice As Integer)

Dim objPais As New ClassPais
Dim colSelecao As Collection

    If Not bTelaFilialEmp Then

        objPais.iCodigo = Codigo_Extrai(Controle("Pais", iIndice).Text)
    
        'Chama a Tela de PaisesLista
        Call Chama_Tela("PaisesLista", colSelecao, objPais, objEventoPais)
        
    End If

End Sub

Public Sub LabelCidade_Click(ByVal iIndice As Integer)

Dim objCidade As New ClassCidades
Dim colSelecao As Collection

    If Not bTelaFilialEmp Then

        objCidade.sDescricao = Controle("Cidade", iIndice).Text
    
        'Chama a Tela de browse
        Call Chama_Tela("CidadeLista", colSelecao, objCidade, objEventoCidade)
        
    End If

End Sub

Private Sub objEventoCidade_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCidade As ClassCidades

On Error GoTo Erro_objEventoCidade_evSelecao

    Set objCidade = obj1

    Controle("Cidade", iTabAtual).Text = objCidade.sDescricao

    objTela.Show

    Exit Sub

Erro_objEventoCidade_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202974)

    End Select

    Exit Sub

End Sub

Public Sub Cidade_Validate(Cancel As Boolean, ByVal iIndice As Integer)

Dim lErro As Long, objCidade As New ClassCidades
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Cidade_Validate

    If Len(Trim(Controle("Cidade", iIndice).Text)) = 0 Then Exit Sub

    objCidade.sDescricao = Controle("Cidade", iIndice).Text
    
    lErro = CF("Cidade_Le_Nome", objCidade)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 202975

    If lErro <> SUCESSO Then gError 202976

    Exit Sub

Erro_Cidade_Validate:

    Cancel = True

    Select Case gErr

        Case 202975

        Case 202976
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CIDADE")
            If vbMsgRes = vbYes Then
                If bTelaFilialEmp Then
                    Call CF("Cidade_Grava", objCidade)
                Else
                    Call Chama_Tela("CidadeCadastro", objCidade)
                End If
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202977)

    End Select

    Exit Sub

End Sub

Public Sub CEP_Validate(Cancel As Boolean, ByVal iIndice As Integer)

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CEP_Validate

    If Len(Trim(Controle("CEP", iIndice).Text)) = 0 Then Exit Sub

    objEndereco.sCEP = Controle("CEP", iIndice).Text
    
    lErro = CF("Endereco_Le_CEP", objEndereco)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 202985
    
    If lErro = SUCESSO Then
        
        vbMsgRes = vbYes
        If Len(Trim(Controle("Logradouro", iIndice).Text)) <> 0 And sCEPAnt(iIndice) <> Controle("CEP", iIndice).Text Then
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUIR_ENDERECO_ATUAL")
        End If
        
        If vbMsgRes = vbYes And sCEPAnt(iIndice) <> Controle("CEP", iIndice).Text Then
        
            Controle("Bairro", iIndice).Text = objEndereco.sBairro
            Controle("Cidade", iIndice).Text = objEndereco.sCidade
            Controle("Estado", iIndice).Text = objEndereco.sSiglaEstado
            Call Estado_Validate(bSGECancelDummy, iIndice)
            Controle("Pais", iIndice).Text = PAIS_BRASIL
            Call Pais_Validate(bSGECancelDummy, iIndice)
            Controle("TipoLogradouro", iIndice).Text = objEndereco.sTipoLogradouro
            Call TipoLogradouro_Validate(bSGECancelDummy, iIndice, True)
            Controle("Logradouro", iIndice).Text = objEndereco.sLogradouro
            
            bMudouCEP = True
        
        End If
    
    End If
    
    sCEPAnt(iIndice) = Controle("CEP", iIndice).Text

    Exit Sub

Erro_CEP_Validate:

    Cancel = True

    Select Case gErr
    
        Case 202985
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202986)

    End Select

    Exit Sub

End Sub

Public Sub CEP_LostFocus(ByVal iIndice As Integer)
    If bMudouCEP Then Call Controle("Cidade", iIndice).SetFocus
End Sub
