VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTelaDadosCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_objUserControl As Object

Dim gobjTiposCamposCust As ClassTiposCamposCust

Const CONTROLE_TIPO_DATA = 0
Const CONTROLE_TIPO_VALOR = 1
Const CONTROLE_TIPO_NUMERO = 2
Const CONTROLE_TIPO_TEXTO = 3

Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

Public Property Get Data(ByVal iIndice As Integer) As Object
     Set Data = objUserControl.Controls("Data").Item(iIndice)
End Property

Public Property Get Numero(ByVal iIndice As Integer) As Object
     Set Numero = objUserControl.Controls("Numero").Item(iIndice)
End Property

Public Property Get Valor(ByVal iIndice As Integer) As Object
     Set Valor = objUserControl.Controls("Valor").Item(iIndice)
End Property

Public Property Get TEXTO(ByVal iIndice As Integer) As Object
     Set TEXTO = objUserControl.Controls("Texto").Item(iIndice)
End Property

Public Property Get Label1(ByVal iIndice As Integer) As Object
     Set Label1 = objUserControl.Controls("Label1").Item(iIndice)
End Property

Public Property Get UpDownData(ByVal iIndice As Integer) As Object
     Set UpDownData = objUserControl.Controls("UpDownData").Item(iIndice)
End Property

Public Property Get Controles() As Object
     Set Controles = objUserControl.Controls("Controles")
End Property

Public Function Move_Tela_Memoria(ByVal objCamposCustomizados As ClassCamposCustomizados, ByVal objTiposCamposCust As ClassTiposCamposCust) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Move_Tela_Memoria

    For iIndice = 1 To CAMPO_CUSTOMIZADO_QTD_REPETICOES

        objCamposCustomizados.dtData(iIndice) = StrParaDate(Data(iIndice).Text)
        objCamposCustomizados.sTexto(iIndice) = TEXTO(iIndice).Text
        objCamposCustomizados.lNumero(iIndice) = StrParaLong(Numero(iIndice).Text)
        objCamposCustomizados.dValor(iIndice) = StrParaDbl(Valor(iIndice).Text)

    Next
    
    lErro = Verifica_Preenchimento(objTiposCamposCust)
    If lErro <> SUCESSO Then gError 181570

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 181570

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181527)

    End Select

    Exit Function

End Function

Public Function Traz_CamposCustomizados_Tela(ByVal objCamposCustomizados As ClassCamposCustomizados) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Traz_CamposCustomizados_Tela

    'Lê o CamposCustomizados que está sendo Passado
    lErro = CF("CamposCustomizados_Le", objCamposCustomizados)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181493

    If lErro = SUCESSO Then

        For iIndice = 1 To CAMPO_CUSTOMIZADO_QTD_REPETICOES

            If objCamposCustomizados.dtData(iIndice) <> DATA_NULA Then
                Data(iIndice).PromptInclude = False
                Data(iIndice).Text = Format(objCamposCustomizados.dtData(iIndice), "dd/mm/yy")
                Data(iIndice).PromptInclude = True
            End If
    
            TEXTO(iIndice).Text = objCamposCustomizados.sTexto(iIndice)
    
            If objCamposCustomizados.lNumero(iIndice) <> 0 Then Numero(iIndice).Text = CStr(objCamposCustomizados.lNumero(iIndice))
            If objCamposCustomizados.dValor(iIndice) <> 0 Then Valor(iIndice).Text = CStr(objCamposCustomizados.dValor(iIndice))

        Next

    End If

    Traz_CamposCustomizados_Tela = SUCESSO

    Exit Function

Erro_Traz_CamposCustomizados_Tela:

    Traz_CamposCustomizados_Tela = gErr

    Select Case gErr

        Case 181493

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181528)

    End Select

    Exit Function

End Function

Public Function UpDownData_DownClick(ByVal iIndice As Integer) As Long

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data(iIndice).SetFocus

    If Len(Data(iIndice).clipText) > 0 Then

        sData = Data(iIndice).Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 181497

        Data(iIndice).Text = sData

    End If
    
    UpDownData_DownClick = SUCESSO

    Exit Function

Erro_UpDownData_DownClick:

    UpDownData_DownClick = gErr

    Select Case gErr

        Case 181497

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181529)

    End Select

    Exit Function

End Function

Public Function UpDownData_UpClick(ByVal iIndice As Integer) As Long

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    Data(iIndice).SetFocus

    If Len(Trim(Data(iIndice).clipText)) > 0 Then

        sData = Data(iIndice).Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 181498

        Data(iIndice).Text = sData

    End If

    UpDownData_UpClick = SUCESSO

    Exit Function

Erro_UpDownData_UpClick:

    UpDownData_UpClick = gErr

    Select Case gErr

        Case 181498

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181530)

    End Select

    Exit Function

End Function

Public Function Data_GotFocus(ByVal iIndice As Integer) As Long
    
    Call MaskEdBox_TrataGotFocus(Data(iIndice), objUserControl.iAlterado)
    
    Data_GotFocus = SUCESSO
    
End Function

Public Function Data_Validate(ByVal iIndice As Integer, Cancel As Boolean) As Long

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data(iIndice).clipText)) <> 0 Then

        lErro = Data_Critica(Data(iIndice).Text)
        If lErro <> SUCESSO Then gError 181499

    End If
    
    Data_Validate = SUCESSO

    Exit Function

Erro_Data_Validate:

    Data_Validate = gErr

    Cancel = True

    Select Case gErr

        Case 181499

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181531)

    End Select

    Exit Function

End Function

Public Function Numero_Validate(ByVal iIndice As Integer, Cancel As Boolean) As Long

Dim lErro As Long

On Error GoTo Erro_Numero_Validate

    'Verifica se Numero1 está preenchida
    If Len(Trim(Numero(iIndice).Text)) <> 0 Then

       'Critica a Numero1
       lErro = Long_Critica(Numero(iIndice).Text)
       If lErro <> SUCESSO Then gError 181512

    End If

    Exit Function

Erro_Numero_Validate:

    Cancel = True

    Select Case gErr

        Case 181512

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181532)

    End Select

    Exit Function

End Function

Public Function Numero_GotFocus(ByVal iIndice As Integer) As Long
    
    Call MaskEdBox_TrataGotFocus(Numero(iIndice), objUserControl.iAlterado)
    
    Numero_GotFocus = SUCESSO
    
End Function

Public Function Valor_Validate(ByVal iIndice As Integer, Cancel As Boolean) As Long

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Verifica se Valor1 está preenchida
    If Len(Trim(Valor(iIndice).Text)) <> 0 Then

       'Critica a Valor1
       lErro = Valor_Positivo_Critica(Valor(iIndice).Text)
       If lErro <> SUCESSO Then gError 181517

    End If
    
    Valor_Validate = SUCESSO

    Exit Function

Erro_Valor_Validate:

    Valor_Validate = gErr

    Cancel = True

    Select Case gErr

        Case 181517

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181533)

    End Select

    Exit Function

End Function

Public Function Valor_GotFocus(ByVal iIndice As Integer) As Long
    
    Call MaskEdBox_TrataGotFocus(Valor(iIndice), objUserControl.iAlterado)
    
    Valor_GotFocus = SUCESSO
    
End Function

Public Function Verifica_Preenchimento(ByVal objTiposCamposCust As ClassTiposCamposCust) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Verifica_Preenchimento

    objTiposCamposCust.sNomeTela = objUserControl.Name

    For iIndice = 1 To CAMPO_CUSTOMIZADO_QTD_REPETICOES

        If Data(iIndice).Enabled Then
            objTiposCamposCust.iDatasPreenchida = iIndice
        End If
        
        If TEXTO(iIndice).Enabled Then
            objTiposCamposCust.iTextosPreenchidos = iIndice
        End If
        
        If Numero(iIndice).Enabled Then
            objTiposCamposCust.iNumerosPreenchidos = iIndice
        End If
        
        If Valor(iIndice).Enabled Then
            objTiposCamposCust.iValoresPreenchidos = iIndice
        End If

    Next

    Verifica_Preenchimento = SUCESSO

    Exit Function

Erro_Verifica_Preenchimento:

    Verifica_Preenchimento = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181535)

    End Select

    Exit Function

End Function

Public Function Exibe_Campos_Customizados() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objTiposCamposCust As New ClassTiposCamposCust

On Error GoTo Erro_Exibe_Campos_Customizados

    objTiposCamposCust.sNomeTela = objUserControl.Name
    
    lErro = CF("TiposCamposCust_Le_NomeTela", objTiposCamposCust)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181577
    
    Set gobjTiposCamposCust = objTiposCamposCust
    
    For iIndice = 1 To objTiposCamposCust.iDatasPreenchida
        Data(iIndice).Visible = True
        UpDownData(iIndice).Visible = True
        Data(iIndice).Enabled = True
        Label1(iIndice + CAMPO_CUSTOMIZADO_VARIACAO_INDEX_DATA).Visible = True
    Next
    
    For iIndice = 1 To objTiposCamposCust.iNumerosPreenchidos
        Numero(iIndice).Visible = True
        Numero(iIndice).Enabled = True
        Label1(iIndice + CAMPO_CUSTOMIZADO_VARIACAO_INDEX_NUMERO).Visible = True
    Next
    
    For iIndice = 1 To objTiposCamposCust.iValoresPreenchidos
        Valor(iIndice).Visible = True
        Valor(iIndice).Enabled = True
        Label1(iIndice + CAMPO_CUSTOMIZADO_VARIACAO_INDEX_VALOR).Visible = True
    Next

    For iIndice = 1 To objTiposCamposCust.iTextosPreenchidos
        TEXTO(iIndice).Visible = True
        TEXTO(iIndice).Enabled = True
        Label1(iIndice + CAMPO_CUSTOMIZADO_VARIACAO_INDEX_TEXTO).Visible = True
    Next
    
    Exibe_Campos_Customizados = SUCESSO

    Exit Function

Erro_Exibe_Campos_Customizados:

    Exibe_Campos_Customizados = gErr

    Select Case gErr

        Case 181577
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSCAMPOSCUST_NAO_CADASTRADO2", gErr, objTiposCamposCust.sNomeTela)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181576)

    End Select

    Exit Function

End Function

Public Function BotaoDadosCustNovo_Click() As Long

Dim lErro As Long

On Error GoTo Erro_BotaoDadosCustNovo_Click

    Select Case Controles.ListIndex
    
        Case CONTROLE_TIPO_DATA
        
            If gobjTiposCamposCust.iDatasPreenchida = CAMPO_CUSTOMIZADO_QTD_REPETICOES Then gError 181579
            
            gobjTiposCamposCust.iDatasPreenchida = gobjTiposCamposCust.iDatasPreenchida + 1
            
            Data(gobjTiposCamposCust.iDatasPreenchida).Visible = True
            Data(gobjTiposCamposCust.iDatasPreenchida).Enabled = True
            UpDownData(gobjTiposCamposCust.iDatasPreenchida).Visible = True
            Label1(gobjTiposCamposCust.iDatasPreenchida + CAMPO_CUSTOMIZADO_VARIACAO_INDEX_DATA).Visible = True
    
        Case CONTROLE_TIPO_VALOR
        
            If gobjTiposCamposCust.iValoresPreenchidos = CAMPO_CUSTOMIZADO_QTD_REPETICOES Then gError 181580
    
            gobjTiposCamposCust.iValoresPreenchidos = gobjTiposCamposCust.iValoresPreenchidos + 1
    
            Valor(gobjTiposCamposCust.iValoresPreenchidos).Visible = True
            Valor(gobjTiposCamposCust.iValoresPreenchidos).Enabled = True
            Label1(gobjTiposCamposCust.iValoresPreenchidos + CAMPO_CUSTOMIZADO_VARIACAO_INDEX_VALOR).Visible = True
    
        Case CONTROLE_TIPO_NUMERO
        
            If gobjTiposCamposCust.iNumerosPreenchidos = CAMPO_CUSTOMIZADO_QTD_REPETICOES Then gError 181581
    
            gobjTiposCamposCust.iNumerosPreenchidos = gobjTiposCamposCust.iNumerosPreenchidos + 1
            
            Numero(gobjTiposCamposCust.iNumerosPreenchidos).Visible = True
            Numero(gobjTiposCamposCust.iNumerosPreenchidos).Enabled = True
            Label1(gobjTiposCamposCust.iNumerosPreenchidos + CAMPO_CUSTOMIZADO_VARIACAO_INDEX_NUMERO).Visible = True
    
        Case CONTROLE_TIPO_TEXTO
        
            If gobjTiposCamposCust.iTextosPreenchidos = CAMPO_CUSTOMIZADO_QTD_REPETICOES Then gError 181582
    
            gobjTiposCamposCust.iTextosPreenchidos = gobjTiposCamposCust.iTextosPreenchidos + 1
    
            TEXTO(gobjTiposCamposCust.iTextosPreenchidos).Visible = True
            TEXTO(gobjTiposCamposCust.iTextosPreenchidos).Enabled = True
            Label1(gobjTiposCamposCust.iTextosPreenchidos + CAMPO_CUSTOMIZADO_VARIACAO_INDEX_TEXTO).Visible = True
    
    End Select

    BotaoDadosCustNovo_Click = SUCESSO
    
    Exit Function
    
Erro_BotaoDadosCustNovo_Click:

    BotaoDadosCustNovo_Click = gErr

    Select Case gErr
    
        Case 181579
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_CAMPOS_DATA", gErr, CAMPO_CUSTOMIZADO_QTD_REPETICOES)
      
        Case 181580
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_CAMPOS_VALOR", gErr, CAMPO_CUSTOMIZADO_QTD_REPETICOES)
      
        Case 181581
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_CAMPOS_NUMERO", gErr, CAMPO_CUSTOMIZADO_QTD_REPETICOES)
      
        Case 181582
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_CAMPOS_TEXTO", gErr, CAMPO_CUSTOMIZADO_QTD_REPETICOES)
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181578)

    End Select

    Exit Function

End Function

Public Function BotaoDadosCustDel_Click() As Long

Dim lErro As Long

On Error GoTo Erro_BotaoDadosCustDel_Click

    Select Case Controles.ListIndex
    
        Case CONTROLE_TIPO_DATA
        
            If gobjTiposCamposCust.iDatasPreenchida = 0 Then gError 181635
                       
            Data(gobjTiposCamposCust.iDatasPreenchida).Visible = False
            Data(gobjTiposCamposCust.iDatasPreenchida).Enabled = False
            UpDownData(gobjTiposCamposCust.iDatasPreenchida).Visible = False
            Label1(gobjTiposCamposCust.iDatasPreenchida + CAMPO_CUSTOMIZADO_VARIACAO_INDEX_DATA).Visible = False
    
            gobjTiposCamposCust.iDatasPreenchida = gobjTiposCamposCust.iDatasPreenchida - 1

        Case CONTROLE_TIPO_VALOR
        
            If gobjTiposCamposCust.iValoresPreenchidos = 0 Then gError 181636
       
            Valor(gobjTiposCamposCust.iValoresPreenchidos).Visible = False
            Valor(gobjTiposCamposCust.iValoresPreenchidos).Enabled = False
            Label1(gobjTiposCamposCust.iValoresPreenchidos + CAMPO_CUSTOMIZADO_VARIACAO_INDEX_VALOR).Visible = False
    
            gobjTiposCamposCust.iValoresPreenchidos = gobjTiposCamposCust.iValoresPreenchidos - 1
   
        Case CONTROLE_TIPO_NUMERO
        
            If gobjTiposCamposCust.iNumerosPreenchidos = 0 Then gError 181637
               
            Numero(gobjTiposCamposCust.iNumerosPreenchidos).Visible = False
            Numero(gobjTiposCamposCust.iNumerosPreenchidos).Enabled = False
            Label1(gobjTiposCamposCust.iNumerosPreenchidos + CAMPO_CUSTOMIZADO_VARIACAO_INDEX_NUMERO).Visible = False
    
            gobjTiposCamposCust.iNumerosPreenchidos = gobjTiposCamposCust.iNumerosPreenchidos - 1
    
        Case CONTROLE_TIPO_TEXTO
        
            If gobjTiposCamposCust.iTextosPreenchidos = 0 Then gError 181638
       
            TEXTO(gobjTiposCamposCust.iTextosPreenchidos).Visible = False
            TEXTO(gobjTiposCamposCust.iTextosPreenchidos).Enabled = False
            Label1(gobjTiposCamposCust.iTextosPreenchidos + CAMPO_CUSTOMIZADO_VARIACAO_INDEX_TEXTO).Visible = False
    
            gobjTiposCamposCust.iTextosPreenchidos = gobjTiposCamposCust.iTextosPreenchidos - 1
    
    End Select

    BotaoDadosCustDel_Click = SUCESSO
    
    Exit Function
    
Erro_BotaoDadosCustDel_Click:

    BotaoDadosCustDel_Click = gErr

    Select Case gErr
    
        Case 181635
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_CAMPO_DATA_VISIVEL", gErr)
      
        Case 181636
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_CAMPO_VALOR_VISIVEL", gErr)
      
        Case 181637
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_CAMPO_NUMERO_VISIVEL", gErr)
      
        Case 181638
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_CAMPO_TEXTO_VISIVEL", gErr)
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181639)

    End Select

    Exit Function

End Function

Private Sub Class_Terminate()
    Set gobjTiposCamposCust = Nothing
    Set m_objUserControl = Nothing
End Sub
