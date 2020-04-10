VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTClientesAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Empresa1(ByVal objCT As Object) As Object
     Set Empresa1 = objCT.objUserControl.Controls("Empresa1")
End Property

Public Property Get MatriculaPrinc(ByVal objCT As Object) As Object
     Set MatriculaPrinc = objCT.objUserControl.Controls("MatriculaPrinc")
End Property

Public Property Get MatriculaPrincDV(ByVal objCT As Object) As Object
     Set MatriculaPrincDV = objCT.objUserControl.Controls("MatriculaPrincDV")
End Property

Public Property Get Empresa2(ByVal objCT As Object) As Object
     Set Empresa2 = objCT.objUserControl.Controls("Empresa2")
End Property

Public Property Get MatriculaSec(ByVal objCT As Object) As Object
     Set MatriculaSec = objCT.objUserControl.Controls("MatriculaSec")
End Property

Public Property Get MatriculaAF(ByVal objCT As Object) As Object
     Set MatriculaAF = objCT.objUserControl.Controls("MatriculaAF")
End Property

Public Property Get MatriculaFRG(ByVal objCT As Object) As Object
     Set MatriculaFRG = objCT.objUserControl.Controls("MatriculaFRG")
End Property

Public Property Get Benemerito(ByVal objCT As Object) As Object
     Set Benemerito = objCT.objUserControl.Controls("Benemerito")
End Property

Public Property Get Fundador(ByVal objCT As Object) As Object
     Set Fundador = objCT.objUserControl.Controls("Fundador")
End Property

Public Property Get RGOrgaoEmissor(ByVal objCT As Object) As Object
     Set RGOrgaoEmissor = objCT.objUserControl.Controls("RGOrgaoEmissor")
End Property

Public Property Get DataEmissaoRG(ByVal objCT As Object) As Object
     Set DataEmissaoRG = objCT.objUserControl.Controls("DataEmissaoRG")
End Property

Public Property Get Sexo(ByVal objCT As Object) As Object
     Set Sexo = objCT.objUserControl.Controls("Sexo")
End Property

Public Property Get DataNasc(ByVal objCT As Object) As Object
     Set DataNasc = objCT.objUserControl.Controls("DataNasc")
End Property

Public Property Get DataFalecimento(ByVal objCT As Object) As Object
     Set DataFalecimento = objCT.objUserControl.Controls("DataFalecimento")
End Property

Public Property Get DataConBenef(ByVal objCT As Object) As Object
     Set DataConBenef = objCT.objUserControl.Controls("DataConBenef")
End Property

Public Property Get Cargo(ByVal objCT As Object) As Object
     Set Cargo = objCT.objUserControl.Controls("Cargo")
End Property

Public Property Get DataInsc(ByVal objCT As Object) As Object
     Set DataInsc = objCT.objUserControl.Controls("DataInsc")
End Property

Public Property Get DataAdmAF(ByVal objCT As Object) As Object
     Set DataAdmAF = objCT.objUserControl.Controls("DataAdmAF")
End Property

Public Property Get TipoApos(ByVal objCT As Object) As Object
     Set TipoApos = objCT.objUserControl.Controls("TipoApos")
End Property

Public Property Get DataAposINSS(ByVal objCT As Object) As Object
     Set DataAposINSS = objCT.objUserControl.Controls("DataAposINSS")
End Property

Public Property Get DataAposFRG(ByVal objCT As Object) As Object
     Set DataAposFRG = objCT.objUserControl.Controls("DataAposFRG")
End Property

Public Property Get ContrSoc(ByVal objCT As Object) As Object
     Set ContrSoc = objCT.objUserControl.Controls("ContrSoc")
End Property

Public Property Get DataUltAtu(ByVal objCT As Object) As Object
     Set DataUltAtu = objCT.objUserControl.Controls("DataUltAtu")
End Property

Public Property Get LocalTrabalho(ByVal objCT As Object) As Object
     Set LocalTrabalho = objCT.objUserControl.Controls("LocalTrabalho")
End Property

Public Property Get NoBenef(ByVal objCT As Object) As Object
     Set NoBenef = objCT.objUserControl.Controls("NoBenef")
End Property

Public Function CLI_Grava_Registro(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CLI_Grava_Registro

    If Len(Trim(Empresa1(objCT).Text)) = 0 Then gError 192959
    If Len(Trim(MatriculaPrinc(objCT).Text)) = 0 Then gError 192960
    If Len(Trim(Sexo(objCT).Text)) = 0 Then gError 192961
    If StrParaDate(DataNasc(objCT).Text) = DATA_NULA Then gError 192962
    If StrParaDate(DataInsc(objCT).Text) = DATA_NULA Then gError 192963
    
    Call Atualiza_NomeReduzido(objCT)
    
    CLI_Grava_Registro = SUCESSO
     
    Exit Function
    
Erro_CLI_Grava_Registro:

    CLI_Grava_Registro = gErr
     
    Select Case gErr
    
        Case 192959
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEAF_EMPRESA1_NAO_PREENCHIDA", gErr)
          
        Case 192960
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEAF_MATPRINC_NAO_PREENCHIDA", gErr)
          
        Case 192961
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEAF_SEXO_NAO_PREENCHIDA", gErr)
          
        Case 192962
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEAF_DATANASC_NAO_PREENCHIDA", gErr)
          
        Case 192963
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEAF_DATAINSC_NAO_PREENCHIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192964)
     
    End Select
     
    Exit Function

End Function

Public Function CLI_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CLI_Form_Load

    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_AF_EMPRESAS, Empresa1(objCT), True, False, True)
    If lErro <> SUCESSO Then gError 192973
    
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_AF_EMPRESAS, Empresa2(objCT), True, False, True)
    If lErro <> SUCESSO Then gError 192974
    
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_AF_TIPOAPOS, TipoApos(objCT), True, False, True)
    If lErro <> SUCESSO Then gError 192975
    
    CLI_Form_Load = SUCESSO
     
    Exit Function
    
Erro_CLI_Form_Load:

    CLI_Form_Load = gErr
     
    Select Case gErr
    
        Case 192973 To 192975

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192976)
     
    End Select
     
    Exit Function

End Function

Public Function CLI_Move_Cliente_Memoria(ByVal objCT As Object, ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim objClienteAF As New ClassClientesAF

On Error GoTo Erro_CLI_Move_Cliente_Memoria

    objClienteAF.dtDataAdmissaoFurnas = StrParaDate(DataAdmAF(objCT).Text)
    objClienteAF.dtDataAposFRG = StrParaDate(DataAposFRG(objCT).Text)
    objClienteAF.dtDataAposINSS = StrParaDate(DataAposINSS(objCT).Text)
    objClienteAF.dtDataConBenf = StrParaDate(DataConBenef(objCT).Text)
    objClienteAF.dtDataExpedicaoRG = StrParaDate(DataEmissaoRG(objCT).Text)
    objClienteAF.dtDataFalecimento = StrParaDate(DataFalecimento(objCT).Text)
    objClienteAF.dtDataInscricao = StrParaDate(DataInsc(objCT).Text)
    objClienteAF.dtDataNascimento = StrParaDate(DataNasc(objCT).Text)
    objClienteAF.dtDataUltAtualizacao = StrParaDate(DataUltAtu(objCT).Text)
    
    If objClienteAF.dtDataFalecimento <> DATA_NULA And objClienteAF.dtDataNascimento > objClienteAF.dtDataFalecimento Then gError 194021
    If objClienteAF.dtDataNascimento > objClienteAF.dtDataInscricao Then gError 194022
    If objClienteAF.dtDataFalecimento <> DATA_NULA And objClienteAF.dtDataInscricao > objClienteAF.dtDataFalecimento Then gError 194023
    
    If Benemerito(objCT).Value = vbChecked Then
        objClienteAF.iBenemerito = MARCADO
    Else
        objClienteAF.iBenemerito = DESMARCADO
    End If
    
    If Empresa1(objCT).ListIndex <> -1 Then objClienteAF.iEmpresa1 = Empresa1(objCT).ItemData(Empresa1(objCT).ListIndex)
    If Empresa2(objCT).ListIndex <> -1 Then objClienteAF.iEmpresa2 = Empresa2(objCT).ItemData(Empresa2(objCT).ListIndex)
    
    If Fundador(objCT).Value = vbChecked Then
        objClienteAF.iFundador = MARCADO
    Else
        objClienteAF.iFundador = DESMARCADO
    End If
    
    'objClienteAF.iStatusAssociado
    If TipoApos(objCT).ListIndex <> -1 Then objClienteAF.iTipoApos = TipoApos(objCT).ItemData(TipoApos(objCT).ListIndex)
    'objClienteAF.iTipoAssociado
    'objClienteAF.lCliente
    objClienteAF.lMatriculaAF = StrParaLong(MatriculaAF(objCT).Text)
    objClienteAF.lMatriculaFRG = StrParaLong(MatriculaFRG(objCT).Text)
    objClienteAF.lMatriculaPrinc = StrParaLong(MatriculaPrinc(objCT).Text)
    objClienteAF.lMatriculaSec = StrParaLong(MatriculaSec(objCT).Text)
    objClienteAF.sCargo = Cargo(objCT).Text
    objClienteAF.sContrSoc = ContrSoc(objCT).Text
    objClienteAF.sLocalTrabalho = LocalTrabalho(objCT).Text
    objClienteAF.sNoBenef = NoBenef(objCT).Text
    'objClienteAF.sObservacao1
    'objClienteAF.sObservacao2
    objClienteAF.sRGOrgaoEmissor = RGOrgaoEmissor(objCT).Text
    objClienteAF.sSexo = Sexo(objCT).Text
    objClienteAF.iMatriculaPrincDV = StrParaInt(MatriculaPrincDV(objCT).Text)

    Set objCliente.objInfoUsu = objClienteAF
    
    Call Atualiza_NomeReduzido(objCT)
    
    CLI_Move_Cliente_Memoria = SUCESSO
     
    Exit Function
    
Erro_CLI_Move_Cliente_Memoria:

    CLI_Move_Cliente_Memoria = gErr
     
    Select Case gErr
    
        Case 194021
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NASC_MAIOR_DATA_FALEC", gErr)
          
        Case 194022
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NASC_MAIOR_DATA_INSC", gErr)
          
        Case 194023
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INSC_MAIOR_DATA_FALEC", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192977)
     
    End Select
     
    Exit Function

End Function

Public Function CLI_Limpa_Tela_Cliente(ByVal objCT As Object) As Long

    Benemerito(objCT).Value = vbUnchecked
    Fundador(objCT).Value = vbUnchecked
    
    Empresa1(objCT).ListIndex = -1
    Empresa2(objCT).ListIndex = -1
    TipoApos(objCT).ListIndex = -1
    Sexo(objCT).ListIndex = -1
    
    CLI_Limpa_Tela_Cliente = SUCESSO
    
End Function

Public Function CLI_Traz_Cliente_Tela(ByVal objCT As Object, ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim objClienteAF As ClassClientesAF

On Error GoTo Erro_CLI_Traz_Cliente_Tela

    lErro = CF("Cliente_Le_Customizado", objCliente)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192978
    
    If Not (objCliente.objInfoUsu Is Nothing) Then
    
        Set objClienteAF = objCliente.objInfoUsu

        DataAdmAF(objCT).PromptInclude = False
        If objClienteAF.dtDataAdmissaoFurnas <> DATA_NULA Then
            DataAdmAF(objCT).Text = Format(objClienteAF.dtDataAdmissaoFurnas, "dd/mm/yyyy")
        Else
            DataAdmAF(objCT).Text = ""
        End If
        DataAdmAF(objCT).PromptInclude = True
        
        DataAposFRG(objCT).PromptInclude = False
        If objClienteAF.dtDataAposFRG <> DATA_NULA Then
            DataAposFRG(objCT).Text = Format(objClienteAF.dtDataAposFRG, "dd/mm/yyyy")
        Else
            DataAposFRG(objCT).Text = ""
        End If
        DataAposFRG(objCT).PromptInclude = True
        
        DataAposINSS(objCT).PromptInclude = False
        If objClienteAF.dtDataAposINSS <> DATA_NULA Then
            DataAposINSS(objCT).Text = Format(objClienteAF.dtDataAposINSS, "dd/mm/yyyy")
        Else
            DataAposINSS(objCT).Text = ""
        End If
        DataAposINSS(objCT).PromptInclude = True
        
        DataConBenef(objCT).PromptInclude = False
        If objClienteAF.dtDataConBenf <> DATA_NULA Then
            DataConBenef(objCT).Text = Format(objClienteAF.dtDataConBenf, "dd/mm/yyyy")
        Else
            DataConBenef(objCT).Text = ""
        End If
        DataConBenef(objCT).PromptInclude = True
        
        DataEmissaoRG(objCT).PromptInclude = False
        If objClienteAF.dtDataExpedicaoRG <> DATA_NULA Then
            DataEmissaoRG(objCT).Text = Format(objClienteAF.dtDataExpedicaoRG, "dd/mm/yyyy")
        Else
            DataEmissaoRG(objCT).Text = ""
        End If
        DataEmissaoRG(objCT).PromptInclude = True
        
        DataFalecimento(objCT).PromptInclude = False
        If objClienteAF.dtDataFalecimento <> DATA_NULA Then
            DataFalecimento(objCT).Text = Format(objClienteAF.dtDataFalecimento, "dd/mm/yyyy")
        Else
            DataFalecimento(objCT).Text = ""
        End If
        DataFalecimento(objCT).PromptInclude = True
        
        DataInsc(objCT).PromptInclude = False
        If objClienteAF.dtDataInscricao <> DATA_NULA Then
            DataInsc(objCT).Text = Format(objClienteAF.dtDataInscricao, "dd/mm/yyyy")
        Else
            DataInsc(objCT).Text = ""
        End If
        DataInsc(objCT).PromptInclude = True
        
        DataNasc(objCT).PromptInclude = False
        If objClienteAF.dtDataNascimento <> DATA_NULA Then
            DataNasc(objCT).Text = Format(objClienteAF.dtDataNascimento, "dd/mm/yyyy")
        Else
            DataNasc(objCT).Text = ""
        End If
        DataNasc(objCT).PromptInclude = True
        
        DataUltAtu(objCT).PromptInclude = False
        If objClienteAF.dtDataUltAtualizacao <> DATA_NULA Then
            DataUltAtu(objCT).Text = Format(objClienteAF.dtDataUltAtualizacao, "dd/mm/yyyy")
        Else
            DataUltAtu(objCT).Text = ""
        End If
        DataUltAtu(objCT).PromptInclude = True
        
        If objClienteAF.iBenemerito = MARCADO Then
            Benemerito(objCT).Value = vbChecked
        Else
            Benemerito(objCT).Value = vbUnchecked
        End If
        
        Call Combo_Seleciona_ItemData(Empresa1(objCT), objClienteAF.iEmpresa1)
        Call Combo_Seleciona_ItemData(Empresa2(objCT), objClienteAF.iEmpresa2)

        If objClienteAF.iFundador = MARCADO Then
            Fundador(objCT).Value = vbChecked
        Else
            Fundador(objCT).Value = vbUnchecked
        End If
        
        'objClienteAF.iStatusAssociado
        Call Combo_Seleciona_ItemData(TipoApos(objCT), objClienteAF.iTipoApos)
        'objClienteAF.iTipoAssociado
        'objClienteAF.lCliente
        
        MatriculaAF(objCT).PromptInclude = False
        If objClienteAF.lMatriculaAF > 0 Then
            MatriculaAF(objCT).Text = CStr(objClienteAF.lMatriculaAF)
        Else
            MatriculaAF(objCT).Text = ""
        End If
        MatriculaAF(objCT).PromptInclude = True
        
        MatriculaFRG(objCT).PromptInclude = False
        If objClienteAF.lMatriculaFRG > 0 Then
            MatriculaFRG(objCT).Text = CStr(objClienteAF.lMatriculaFRG)
        Else
            MatriculaFRG(objCT).Text = ""
        End If
        MatriculaFRG(objCT).PromptInclude = True
        
        MatriculaPrinc(objCT).PromptInclude = False
        If objClienteAF.lMatriculaPrinc > 0 Then
            MatriculaPrinc(objCT).Text = CStr(objClienteAF.lMatriculaPrinc)
        Else
            MatriculaPrinc(objCT).Text = ""
        End If
        MatriculaPrinc(objCT).PromptInclude = True
        
        MatriculaPrincDV(objCT).PromptInclude = False
        If objClienteAF.iMatriculaPrincDV > 0 Then
            MatriculaPrincDV(objCT).Text = CStr(objClienteAF.iMatriculaPrincDV)
        Else
            MatriculaPrincDV(objCT).Text = ""
        End If
        MatriculaPrincDV(objCT).PromptInclude = True
        
        MatriculaSec(objCT).PromptInclude = False
        If objClienteAF.lMatriculaSec > 0 Then
            MatriculaSec(objCT).Text = CStr(objClienteAF.lMatriculaSec)
        Else
            MatriculaSec(objCT).Text = ""
        End If
        MatriculaSec(objCT).PromptInclude = True
        
        Cargo(objCT).Text = objClienteAF.sCargo
        ContrSoc(objCT).Text = objClienteAF.sContrSoc
        LocalTrabalho(objCT).Text = objClienteAF.sLocalTrabalho
        NoBenef(objCT).Text = objClienteAF.sNoBenef
        'objClienteAF.sObservacao1
        'objClienteAF.sObservacao2
        RGOrgaoEmissor(objCT).Text = objClienteAF.sRGOrgaoEmissor
        
        Call CF("SCombo_Seleciona2", Sexo(objCT), objClienteAF.sSexo)
        
    End If
    
    CLI_Traz_Cliente_Tela = SUCESSO
     
    Exit Function
    
Erro_CLI_Traz_Cliente_Tela:

    CLI_Traz_Cliente_Tela = gErr
     
    Select Case gErr
    
        Case 192978
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192979)
     
    End Select
     
    Exit Function

End Function

Private Sub UpDown_DownClick(ByVal objMaskData As Object)

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_DownClick_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(objMaskData, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 192980

    Exit Sub

Erro_DownClick_DownClick:

    Select Case gErr

        Case 192981

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192982)

    End Select

    Exit Sub

End Sub

Private Sub UpDown_UpClick(ByVal objMaskData As Object)

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_DownClick_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(objMaskData, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 192983

    Exit Sub

Erro_DownClick_UpClick:

    Select Case gErr

        Case 192983

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192984)

    End Select

    Exit Sub

End Sub

Public Sub Empresa1_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Empresa1_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Empresa2_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Empresa2_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MatriculaPrinc_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MatriculaPrinc_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(MatriculaPrinc(objCT), objCT.iAlterado)
End Sub

Public Sub MatriculaPrinc_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MatriculaPrinc_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(MatriculaPrinc(objCT).Text)) <> 0 Then

        'Critica se é um Long
        lErro = Long_Critica(MatriculaPrinc(objCT).Text)
        If lErro <> SUCESSO Then gError 192985
        
    End If

    Call Atualiza_NomeReduzido(objCT)
    
    Exit Sub

Erro_MatriculaPrinc_Validate:

    Cancel = True

    Select Case gErr

        Case 192985

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192986)

    End Select

    Exit Sub

End Sub

Public Sub MatriculaPrincDV_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MatriculaPrincDV_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(MatriculaPrincDV(objCT), objCT.iAlterado)
End Sub

Public Sub MatriculaPrincDV_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MatriculaPrincDV_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(MatriculaPrincDV(objCT).Text)) <> 0 Then

'        'Critica se é um Long
'        lErro = Inteiro_Critica(MatriculaPrincDV(objCT).Text)
'        If lErro <> SUCESSO Then gError 192985
        
    End If

    Exit Sub

Erro_MatriculaPrincDV_Validate:

    Cancel = True

    Select Case gErr

        Case 192985

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192986)

    End Select

    Exit Sub

End Sub

Public Sub MatriculaSec_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MatriculaSec_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(MatriculaSec(objCT), objCT.iAlterado)
End Sub

Public Sub MatriculaSec_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MatriculaSec_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(MatriculaSec(objCT).Text)) <> 0 Then

        'Critica se é um Long
        lErro = Long_Critica(MatriculaSec(objCT).Text)
        If lErro <> SUCESSO Then gError 192987
        
    End If

    Exit Sub

Erro_MatriculaSec_Validate:

    Cancel = True

    Select Case gErr

        Case 192987

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192988)

    End Select

    Exit Sub

End Sub

Public Sub MatriculaAF_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MatriculaAF_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(MatriculaAF(objCT), objCT.iAlterado)
End Sub

Public Sub MatriculaAF_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MatriculaAF_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(MatriculaAF(objCT).Text)) <> 0 Then

        'Critica se é um Long
        lErro = Long_Critica(MatriculaAF(objCT).Text)
        If lErro <> SUCESSO Then gError 192989
        
    End If

    Exit Sub

Erro_MatriculaAF_Validate:

    Cancel = True

    Select Case gErr

        Case 192989

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192990)

    End Select

    Exit Sub

End Sub

Public Sub MatriculaFRG_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MatriculaFRG_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(MatriculaFRG(objCT), objCT.iAlterado)
End Sub

Public Sub MatriculaFRG_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MatriculaFRG_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(MatriculaFRG(objCT).Text)) <> 0 Then

        'Critica se é um Long
        lErro = Long_Critica(MatriculaFRG(objCT).Text)
        If lErro <> SUCESSO Then gError 192991
        
    End If

    Exit Sub

Erro_MatriculaFRG_Validate:

    Cancel = True

    Select Case gErr

        Case 192991

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192992)

    End Select

    Exit Sub

End Sub

Public Sub Benemerito_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Fundador_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub RGOrgaoEmissor_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataEmissaoRG_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataEmissaoRG_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(DataEmissaoRG(objCT), objCT.iAlterado)
End Sub

Public Sub DataEmissaoRG_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoRG_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(DataEmissaoRG(objCT).clipText)) <> 0 Then

        'Critica se é um Long
        lErro = Data_Critica(DataEmissaoRG(objCT).Text)
        If lErro <> SUCESSO Then gError 192993
        
    End If

    Exit Sub

Erro_DataEmissaoRG_Validate:

    Cancel = True

    Select Case gErr

        Case 192993

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192994)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoRG_UpClick(ByVal objCT As Object)
    Call UpDown_UpClick(DataEmissaoRG(objCT))
End Sub

Public Sub UpDownEmissaoRG_DownClick(ByVal objCT As Object)
    Call UpDown_DownClick(DataEmissaoRG(objCT))
End Sub

Public Sub DataNasc_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataNasc_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(DataNasc(objCT), objCT.iAlterado)
End Sub

Public Sub DataNasc_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataNasc_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(DataNasc(objCT).clipText)) <> 0 Then

        'Critica se é um Long
        lErro = Data_Critica(DataNasc(objCT).Text)
        If lErro <> SUCESSO Then gError 192995
        
    End If

    Call Atualiza_NomeReduzido(objCT)
    
    Exit Sub

Erro_DataNasc_Validate:

    Cancel = True

    Select Case gErr

        Case 192995

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192996)

    End Select

    Exit Sub

End Sub

Public Sub UpDownNasc_UpClick(ByVal objCT As Object)
    Call UpDown_UpClick(DataNasc(objCT))
End Sub

Public Sub UpDownNasc_DownClick(ByVal objCT As Object)
    Call UpDown_DownClick(DataNasc(objCT))
End Sub

Public Sub DataFalecimento_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataFalecimento_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(DataFalecimento(objCT), objCT.iAlterado)
End Sub

Public Sub DataFalecimento_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFalecimento_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(DataFalecimento(objCT).clipText)) <> 0 Then

        'Critica se é um Long
        lErro = Data_Critica(DataFalecimento(objCT).Text)
        If lErro <> SUCESSO Then gError 192997
        
    End If

    Exit Sub

Erro_DataFalecimento_Validate:

    Cancel = True

    Select Case gErr

        Case 192997

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192998)

    End Select

    Exit Sub

End Sub

Public Sub UpDownFalecimento_UpClick(ByVal objCT As Object)
    Call UpDown_UpClick(DataFalecimento(objCT))
End Sub

Public Sub UpDownFalecimento_DownClick(ByVal objCT As Object)
    Call UpDown_DownClick(DataFalecimento(objCT))
End Sub

Public Sub DataInsc_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataInsc_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(DataInsc(objCT), objCT.iAlterado)
End Sub

Public Sub DataInsc_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInsc_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(DataInsc(objCT).clipText)) <> 0 Then

        'Critica se é um Long
        lErro = Data_Critica(DataInsc(objCT).Text)
        If lErro <> SUCESSO Then gError 192999
        
    End If

    Call Atualiza_NomeReduzido(objCT)
    
    Exit Sub

Erro_DataInsc_Validate:

    Cancel = True

    Select Case gErr

        Case 192999

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194000)

    End Select

    Exit Sub

End Sub

Public Sub UpDownInsc_UpClick(ByVal objCT As Object)
    Call UpDown_UpClick(DataInsc(objCT))
End Sub

Public Sub UpDownInsc_DownClick(ByVal objCT As Object)
    Call UpDown_DownClick(DataInsc(objCT))
End Sub

Public Sub DataAdmAF_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataAdmAF_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(DataAdmAF(objCT), objCT.iAlterado)
End Sub

Public Sub DataAdmAF_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAdmAF_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(DataAdmAF(objCT).clipText)) <> 0 Then

        'Critica se é um Long
        lErro = Data_Critica(DataAdmAF(objCT).Text)
        If lErro <> SUCESSO Then gError 194001
        
    End If

    Exit Sub

Erro_DataAdmAF_Validate:

    Cancel = True

    Select Case gErr

        Case 194001

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194002)

    End Select

    Exit Sub

End Sub

Public Sub UpDownAdmAF_UpClick(ByVal objCT As Object)
    Call UpDown_UpClick(DataAdmAF(objCT))
End Sub

Public Sub UpDownAdmAF_DownClick(ByVal objCT As Object)
    Call UpDown_DownClick(DataAdmAF(objCT))
End Sub

Public Sub DataAposINSS_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataAposINSS_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(DataAposINSS(objCT), objCT.iAlterado)
End Sub

Public Sub DataAposINSS_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAposINSS_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(DataAposINSS(objCT).clipText)) <> 0 Then

        'Critica se é um Long
        lErro = Data_Critica(DataAposINSS(objCT).Text)
        If lErro <> SUCESSO Then gError 194003
        
    End If

    Exit Sub

Erro_DataAposINSS_Validate:

    Cancel = True

    Select Case gErr

        Case 194003

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194004)

    End Select

    Exit Sub

End Sub

Public Sub UpDownAposINSS_UpClick(ByVal objCT As Object)
    Call UpDown_UpClick(DataAposINSS(objCT))
End Sub

Public Sub UpDownAposINSS_DownClick(ByVal objCT As Object)
    Call UpDown_DownClick(DataAposINSS(objCT))
End Sub

Public Sub DataAposFRG_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataAposFRG_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(DataAposFRG(objCT), objCT.iAlterado)
End Sub

Public Sub DataAposFRG_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAposFRG_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(DataAposFRG(objCT).clipText)) <> 0 Then

        'Critica se é um Long
        lErro = Data_Critica(DataAposFRG(objCT).Text)
        If lErro <> SUCESSO Then gError 194005
        
    End If

    Exit Sub

Erro_DataAposFRG_Validate:

    Cancel = True

    Select Case gErr

        Case 194005

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194006)

    End Select

    Exit Sub

End Sub

Public Sub UpDownAposFRG_UpClick(ByVal objCT As Object)
    Call UpDown_UpClick(DataAposFRG(objCT))
End Sub

Public Sub UpDownAposFRG_DownClick(ByVal objCT As Object)
    Call UpDown_DownClick(DataAposFRG(objCT))
End Sub

Public Sub DataUltAtu_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataUltAtu_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(DataUltAtu(objCT), objCT.iAlterado)
End Sub

Public Sub DataUltAtu_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataUltAtu_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(DataUltAtu(objCT).clipText)) <> 0 Then

        'Critica se é um Long
        lErro = Data_Critica(DataUltAtu(objCT).Text)
        If lErro <> SUCESSO Then gError 194007
        
    End If

    Exit Sub

Erro_DataUltAtu_Validate:

    Cancel = True

    Select Case gErr

        Case 194007

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194008)

    End Select

    Exit Sub

End Sub

Public Sub UpDownUltAtu_UpClick(ByVal objCT As Object)
    Call UpDown_UpClick(DataUltAtu(objCT))
End Sub

Public Sub UpDownUltAtu_DownClick(ByVal objCT As Object)
    Call UpDown_DownClick(DataUltAtu(objCT))
End Sub

Public Sub Sexo_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Sexo_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Cargo_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoApos_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoApos_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ContrSoc_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NoBenef_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub LocalTrabalho_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataConBenef_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataConBenef_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(DataConBenef(objCT), objCT.iAlterado)
End Sub

Public Sub DataConBenef_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataConBenef_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(DataConBenef(objCT).clipText)) <> 0 Then

        'Critica se é um Long
        lErro = Data_Critica(DataConBenef(objCT).Text)
        If lErro <> SUCESSO Then gError 194009
        
    End If

    Exit Sub

Erro_DataConBenef_Validate:

    Cancel = True

    Select Case gErr

        Case 194009

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194010)

    End Select

    Exit Sub

End Sub

Public Sub UpDownConBenef_UpClick(ByVal objCT As Object)
    Call UpDown_UpClick(DataConBenef(objCT))
End Sub

Public Sub UpDownConBenef_DownClick(ByVal objCT As Object)
    Call UpDown_DownClick(DataConBenef(objCT))
End Sub

Private Sub Atualiza_NomeReduzido(ByVal objCT As Object)
        
Dim sNomeReduzido As String, iFimNome As Integer, sEmp1 As String

    sEmp1 = "F"
    If Empresa1(objCT).ListIndex <> -1 Then
        
        Select Case Empresa1(objCT).ItemData(Empresa1(objCT).ListIndex)
        
            Case 2
                sEmp1 = "E"
            
            Case 4
                sEmp1 = "R"
            
        End Select
    
    End If
    
    sNomeReduzido = sEmp1 & Trim(MatriculaPrinc(objCT).Text) & "." & Trim(objCT.RazaoSocial)
    iFimNome = InStr(1, Trim(sNomeReduzido), " ")
    If iFimNome > 1 Then
        sNomeReduzido = Left(sNomeReduzido, iFimNome - 1)
    End If
    objCT.NomeReduzido = Trim(Left(sNomeReduzido, STRING_CLIENTE_NOME_REDUZIDO))

End Sub
