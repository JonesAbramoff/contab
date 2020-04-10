VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmChamaFuncoes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

''Function E_Zero(dNumero As Double) As Boolean
'''Retorna True se dNumero for zero quando arredondado até precisão de FORMATO_ESTOQUE
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_E_Zero
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_E_ZERO", X)
''    If lErro <> SUCESSO Then Error 25666
''
''    E_Zero = X.E_Zero(dNumero)
''
''    Exit Function
''
''Erro_E_Zero:
''
''    E_Zero = False
''
''    Select Case Err
''
''        Case 25666
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142076)
''
''    End Select
''
''End Function
''
''Function IniciaLetra(sString As String) As Boolean
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_IniciaLetra
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_INICIA_LETRA", X)
''    If lErro <> SUCESSO Then Error 25003
''
''    IniciaLetra = X.IniciaLetra(sString)
''
''    Exit Function
''
''Erro_IniciaLetra:
''
''    IniciaLetra = False
''
''    Select Case Err
''
''        Case 25003
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142077)
''
''    End Select
''
''End Function
''
''Function MesNome(iMes As Integer, sMes As String) As Long
'''Entra com o número do mês. Retorna-o por extenso.
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_MesNome
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_MES_NOME", X)
''    If lErro <> SUCESSO Then Error 25211
''
''    MesNome = X.MesNome(iMes, sMes)
''    Exit Function
''
''Erro_MesNome:
''
''    MesNome = Err
''
''    Select Case Err
''
''        Case 25211
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142078)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function MesNumero(sMes As String, iMes As Integer) As Long
'''Entra com o mês por extenso. Retorna-o como um inteiro.
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_MesNumero
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_MES_NUMERO", X)
''    If lErro <> SUCESSO Then Error 25210
''
''    MesNumero = X.MesNumero(sMes, iMes)
''    Exit Function
''
''Erro_MesNumero:
''
''    MesNumero = Err
''
''    Select Case Err
''
''        Case 25210
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142079)
''
''    End Select
''
''    Exit Function
''
''End Function
''Function Inteiro_Critica(sNumero As String) As Long
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Inteiro_Critica
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_INTEIRO_CRITICA", X)
''    If lErro <> SUCESSO Then Error 6652
''
''    Inteiro_Critica = X.Inteiro_Critica(sNumero)
''    Exit Function
''
''Erro_Inteiro_Critica:
''
''    Inteiro_Critica = Err
''
''    Select Case Err
''
''        Case 6652
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142080)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Long_Critica(sNumero As String) As Long
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Long_Critica
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_LONG_CRITICA", X)
''    If lErro <> SUCESSO Then Error 6653
''
''    Long_Critica = X.Long_Critica(sNumero)
''
''    Exit Function
''
''Erro_Long_Critica:
''
''    Long_Critica = Err
''
''    Select Case Err
''
''        Case 6653
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142081)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Porcentagem_Critica(sString As String) As Long
'''Critica o valor passado em sString , se nao for entre 0 e 100 retorna erro.
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Porcentagem_Critica
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_PORCENTAGEM_CRITICA", X)
''    If lErro <> SUCESSO Then Error 6654
''
''    Porcentagem_Critica = X.Porcentagem_Critica(sString)
''    Exit Function
''
''Erro_Porcentagem_Critica:
''
''    Porcentagem_Critica = Err
''
''    Select Case Err
''
''        Case 6654
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142082)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Data_Critica(ByVal vDataText As Variant) As Long
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Data_Critica
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_DATA_CRITICA", X)
''    If lErro <> SUCESSO Then Error 7002
''
''    Data_Critica = X.Data_Critica(vDataText)
''    Exit Function
''
''Erro_Data_Critica:
''
''    Data_Critica = Err
''
''    Select Case Err
''
''        Case 7002
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142083)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Cpf_Critica(ByVal sCpf As String) As Long
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Cpf_Critica
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_CPF_CRITICA", X)
''    If lErro <> SUCESSO Then Error 6635
''
''
''    Cpf_Critica = X.Cpf_Critica(sCpf)
''    Exit Function
''
''Erro_Cpf_Critica:
''
''    Cpf_Critica = Err
''
''    Select Case Err
''
''        Case 6635
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142084)
''
''    End Select
''
''    Exit Function
''
''End Function

''Function Cgc_Critica(ByVal sCgc As String) As Long
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Cgc_Critica
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_CGC_CRITICA", X)
''    If lErro <> SUCESSO Then Error 6645
''
''
''    Cgc_Critica = X.Cgc_Critica(sCgc)
''    Exit Function
''
''Erro_Cgc_Critica:
''
''    Cgc_Critica = Err
''
''    Select Case Err
''
''        Case 6645
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142085)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Codigo_Extrai(sString As String) As Integer
'''Extrai codigo inteiro de uma string que tem codigo concatenado
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Codigo_Extrai
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_CODIGO_EXTRAI", X)
''    If lErro <> SUCESSO Then Error 6648
''
''
''    Codigo_Extrai = X.Codigo_Extrai(sString)
''    Exit Function
''
''Erro_Codigo_Extrai:
''
''    Codigo_Extrai = Err
''
''    Select Case Err
''
''        Case 6648
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142086)
''
''    End Select
''
''    Exit Function
''
''
''End Function
''
''Function LCodigo_Extrai(sString As String) As Long
'''Extrai codigo long de uma string que tem codigo concatenado
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_LCodigo_Extrai
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_LCODIGO_EXTRAI", X)
''    If lErro <> SUCESSO Then Error 6649
''
''
''    LCodigo_Extrai = X.LCodigo_Extrai(sString)
''    Exit Function
''
''Erro_LCodigo_Extrai:
''
''    LCodigo_Extrai = Err
''
''    Select Case Err
''
''        Case 6649
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142087)
''
''    End Select
''
''    Exit Function
''
''
''End Function
''
''Function Combo_Item_Igual(objComboBox As Object) As Long
'''Seleciona ítem na ComboBox referente ao objComboBox.text passado por parametro
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Combo_Item_Igual
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_COMBO_ITEM_IGUAL", X)
''    If lErro <> SUCESSO Then Error 6650
''
''
''    Combo_Item_Igual = X.Combo_Item_Igual(objComboBox)
''
''    Exit Function
''
''Erro_Combo_Item_Igual:
''
''    Combo_Item_Igual = Err
''
''    Select Case Err
''
''        Case 6650
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142088)
''
''    End Select
''
''    Exit Function
''
''
''End Function
''
''Function Combo_Seleciona(objComboBox As Object, iCodigo As Integer) As Long
'''Seleciona o ítem da Combo correspondente ao Text que pode ser numérico ou string
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Combo_Seleciona
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_COMBO_SELECIONA", X)
''    If lErro <> SUCESSO Then Error 6740
''
''
''    Combo_Seleciona = X.Combo_Seleciona(objComboBox, iCodigo)
''
''    Exit Function
''
''Erro_Combo_Seleciona:
''
''    Combo_Seleciona = Err
''
''    Select Case Err
''
''        Case 6740
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142089)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Combo_Seleciona_Grid(objComboBox As Object, iCodigo As Integer) As Long
'''Seleciona o ítem da Combo correspondente ao Text que pode ser numérico ou string
'''Para ser usada em Grids no lugar de Combo_Seleciona.
'''Não segura foco em caso de erro.
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Combo_Seleciona_Grid
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_COMBO_SELECIONA_GRID", X)
''    If lErro <> SUCESSO Then Error 25087
''
''
''    Combo_Seleciona_Grid = X.Combo_Seleciona_Grid(objComboBox, iCodigo)
''    Exit Function
''
''Erro_Combo_Seleciona_Grid:
''
''    Combo_Seleciona_Grid = Err
''
''    Select Case Err
''
''        Case 25087
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142090)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Combo_Item_Seleciona(objComboBox As Object) As Long
'''Seleciona ítem na ComboBox referente ao codigo ( = 4 primeiros caracteres de objComboBox.text) passado por parametro
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Combo_Item_Seleciona
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_COMBO_ITEM_SELECIONA", X)
''    If lErro <> SUCESSO Then Error 6651
''
''
''    Combo_Item_Seleciona = X.Combo_Item_Seleciona(objComboBox)
''
''    Exit Function
''
''Erro_Combo_Item_Seleciona:
''
''    Combo_Item_Seleciona = Err
''
''    Select Case Err
''
''        Case 6651
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142091)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Data_Aumenta(sData As String) As Long
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Data_Aumenta
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_DATA_AUMENTA", X)
''    If lErro <> SUCESSO Then Error 5923
''
''
''    Data_Aumenta = X.Data_Aumenta(sData)
''    Exit Function
''
''Erro_Data_Aumenta:
''
''    Data_Aumenta = Err
''
''    Select Case Err
''
''        Case 5923
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142092)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Data_Diminui(sData As String) As Long
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Data_Diminui
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_DATA_DIMINUI", X)
''    If lErro <> SUCESSO Then Error 5924
''
''
''    Data_Diminui = X.Data_Diminui(sData)
''
''    Exit Function
''
''Erro_Data_Diminui:
''
''    Data_Diminui = Err
''
''    Select Case Err
''
''        Case 5924
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142093)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Data_Up_Down_Click(MaskEdData As Object, ByVal iOpcao As Integer) As Long
'''recebe o MaskEdBox de uma Data e a opção ( AUMENTA_DATA ou DIMINUI_DATA )
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Data_Up_Down_Click
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_DATA_UP_DOWN_CLICK", X)
''    If lErro <> SUCESSO Then Error 14081
''
''
''    Data_Up_Down_Click = X.Data_Up_Down_Click(MaskEdData, iOpcao)
''
''    Exit Function
''
''Erro_Data_Up_Down_Click:
''
''    Data_Up_Down_Click = Err
''
''    Select Case Err
''
''        Case 14081
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142094)
''
''    End Select
''
''    Exit Function
''
''End Function

Function BrowseUsuarioCampo_Grava(colBrowseUsuarioCampo As Collection) As Long
'exclui as informacoes antigas para o usuario x tela em questao e grava as novas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BrowseUsuarioCampo_Grava
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSEUSUARIOCAMPO_GRAVA", X)
    If lErro <> SUCESSO Then Error 9101
    
    BrowseUsuarioCampo_Grava = X.BrowseUsuarioCampo_Grava(colBrowseUsuarioCampo)
    
    Exit Function
    
Erro_BrowseUsuarioCampo_Grava:

    BrowseUsuarioCampo_Grava = Err
    
    Select Case Err

        Case 9101
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142095)
        
    End Select
    
    Exit Function

End Function

Function BrowseUsuarioOrdenacao_Grava(objBrowseUsuarioOrdenacao As AdmBrowseUsuarioOrdenacao) As Long
'grava a informacao de ordenacao para o usuario x tela em questão

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BrowseUsuarioOrdenacao_Grava
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSEUSUARIOORDENACAO_GRAVA", X)
    If lErro <> SUCESSO Then Error 9110
    
    BrowseUsuarioOrdenacao_Grava = X.BrowseUsuarioOrdenacao_Grava(objBrowseUsuarioOrdenacao)
    
    Exit Function
    
Erro_BrowseUsuarioOrdenacao_Grava:

    BrowseUsuarioOrdenacao_Grava = Err
    
    Select Case Err

        Case 9110
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142096)
        
    End Select
    
    Exit Function

End Function

Function BrowseUsuarioCampo_Le(sNomeTela As String, sCodUsuario As String, colBrowseUsuarioCampo As Collection) As Long
'le todos os campos da tela de browse para o usuario especificado e coloca os resultados na coleção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BrowseUsuarioCampo_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSEUSUARIOCAMPO_LE", X)
    If lErro <> SUCESSO Then Error 9035
    
    
    BrowseUsuarioCampo_Le = X.BrowseUsuarioCampo_Le(sNomeTela, sCodUsuario, colBrowseUsuarioCampo)
    
    Exit Function
    
Erro_BrowseUsuarioCampo_Le:

    BrowseUsuarioCampo_Le = Err
    
    Select Case Err

        Case 9035
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142097)
        
    End Select
    
    Exit Function

End Function

Function BrowseUsuarioOrdenacao_Le(objBrowseUsuarioOrdenacao As AdmBrowseUsuarioOrdenacao) As Long
'le a ordenacao escolhida pelo usuario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BrowseUsuarioOrdenacao_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSEUSUARIOORDENACAO_LE", X)
    If lErro <> SUCESSO Then Error 9045
    
    
    BrowseUsuarioOrdenacao_Le = X.BrowseUsuarioOrdenacao_Le(objBrowseUsuarioOrdenacao)
    
    Exit Function
    
Erro_BrowseUsuarioOrdenacao_Le:

    BrowseUsuarioOrdenacao_Le = Err
    
    Select Case Err

        Case 9045
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142098)
        
    End Select
    
    Exit Function

End Function

Function BrowseIndice_Le(sNomeTela As String, colBrowseIndice As Collection) As Long
'le todos os indices referentes a tela especificada e coloca os resultados na coleção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BrowseIndice_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSEINDICE_LE", X)
    If lErro <> SUCESSO Then Error 9052
    
    
    BrowseIndice_Le = X.BrowseIndice_Le(sNomeTela, colBrowseIndice)
    
    Exit Function
    
Erro_BrowseIndice_Le:

    BrowseIndice_Le = Err
    
    Select Case Err

        Case 9052
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142099)
        
    End Select
    
    Exit Function

End Function

Function Browse_Prepara_Comando_SQL(sSQL As String, iNumCampo As Integer, lComando As Long, avCampo() As Variant, colSelecao As Collection, iBind As Integer) As Long
'prepara o comando SQL passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Browse_Prepara_Comando_SQL
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSE_PREPARA_COMANDO_SQL", X)
    If lErro <> SUCESSO Then Error 9066
    
    
    Browse_Prepara_Comando_SQL = X.Browse_Prepara_Comando_SQL(sSQL, iNumCampo, lComando, avCampo, colSelecao, iBind)
    
    Exit Function
    
Erro_Browse_Prepara_Comando_SQL:

    Browse_Prepara_Comando_SQL = Err
    
    Select Case Err

        Case 9066
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142100)
        
    End Select
    
    Exit Function

End Function

Function Campos_Le(objCampo As AdmCampos) As Long
'le a ordenacao escolhida pelo usuario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Campos_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CAMPOS_LE", X)
    If lErro <> SUCESSO Then Error 9062
    
    
    Campos_Le = X.Campos_Le(objCampo)
    
    Exit Function
    
Erro_Campos_Le:

    Campos_Le = Err
    
    Select Case Err

        Case 9062
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142101)
        
    End Select
    
    Exit Function

End Function

Function Campos_Le1(objCampo As AdmCampos, lComando As Long) As Long
'le o titulo do grid do campo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Campos_Le1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CAMPOS_LE1", X)
    If lErro <> SUCESSO Then Error 9185
    
    
    Campos_Le1 = X.Campos_Le1(objCampo, lComando)
    
    Exit Function
    
Erro_Campos_Le1:

    Campos_Le1 = Err
    
    Select Case Err

        Case 9185
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142102)
        
    End Select
    
    Exit Function

End Function

Function Campos_Le2(objCampo As AdmCampos, lComando As Long) As Long
'le o titulo do grid do campo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Campos_Le2
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CAMPOS_LE2", X)
    If lErro <> SUCESSO Then Error 9186
    
    
    Campos_Le2 = X.Campos_Le2(objCampo, lComando)
    
    Exit Function
    
Erro_Campos_Le2:

    Campos_Le2 = Err
    
    Select Case Err

        Case 9186
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142103)
        
    End Select
    
    Exit Function

End Function

Function Browse_Le_NumRegistros(lNumReg As Long, lComando As Long, sSQL As String, avCampo() As Variant) As Long
'retorna o número de registros contidos na tabela sNomeTabela

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Browse_Le_NumRegistros
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSE_LE_NUMREGISTROS", X)
    If lErro <> SUCESSO Then Error 9067
    
    
    Browse_Le_NumRegistros = X.Browse_Le_NumRegistros(lNumReg, lComando, sSQL, avCampo)
    
    Exit Function
        
Erro_Browse_Le_NumRegistros:

    Browse_Le_NumRegistros = Err
    
    Select Case Err

        Case 9067
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142104)
        
    End Select
    
    Exit Function

End Function

Function Browse_Le_NumRegistros1(sSQL As String, lNumReg As Long, colSegmentos As Collection, colSelecao As Collection, objBrowse As AdmBrowse) As Long
'retorna o número de registros selecionados pelo comando SQL passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Browse_Le_NumRegistros1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSE_LE_NUMREGISTROS1", X)
    If lErro <> SUCESSO Then Error 9141
    
    
    Browse_Le_NumRegistros1 = X.Browse_Le_NumRegistros1(sSQL, lNumReg, colSegmentos, colSelecao, objBrowse)
    
    Exit Function
    
Erro_Browse_Le_NumRegistros1:

    Browse_Le_NumRegistros1 = Err
    
    Select Case Err

        Case 9141
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142105)
        
    End Select
    
    Exit Function

End Function

Function Browse_Le_Posicionado(ByVal colReg As Collection, ByVal lPosicao As Long, ByVal lNumReg As Long, ByVal lComando As Long, ByVal colValorCampo As Collection, sSQL As String, avCampo() As Variant) As Long
'le os registros lNumReg a partir da posicao lPosicao e coloca-os e colReg.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Browse_Le_Posicionado
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSE_LE_POSICIONADO", X)
    If lErro <> SUCESSO Then Error 9068
    
    
    Browse_Le_Posicionado = X.Browse_Le_Posicionado(colReg, lPosicao, lNumReg, lComando, colValorCampo, sSQL, avCampo)
    
    Exit Function
    
Erro_Browse_Le_Posicionado:

    Browse_Le_Posicionado = Err
    
    Select Case Err

        Case 9068
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142106)
        
    End Select
    
    Exit Function

End Function

Function Codigos_Le(sTabela As String, sCampo_Codigo As String, iCodigo_Tipo As enumTipoCodigo, colCodigo As Collection, Optional iTamanho_Nome As Integer) As Long
'Le todos os codigos do campo sCampo_Codigo de sTabela , de acordo com o tipo (iCodigo_Tipo) e devolve na colecao colCodigo.
'Se iCodigoTipo for TIPO_STR tem que vir preenchido iTamanho_Nome

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Codigos_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CODIGOS_LE", X)
    If lErro <> SUCESSO Then Error 6647
    
    
    
    If IsMissing(iTamanho_Nome) Then
        Codigos_Le = X.Codigos_Le(sTabela, sCampo_Codigo, iCodigo_Tipo, colCodigo)
    Else
        Codigos_Le = X.Codigos_Le(sTabela, sCampo_Codigo, iCodigo_Tipo, colCodigo, iTamanho_Nome)
    End If
    
    Exit Function
    
Erro_Codigos_Le:

    Codigos_Le = Err
    
    Select Case Err

        Case 6647
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142107)
        
    End Select
    
    Exit Function

End Function

Function Config_ObterNumInt(sTabelaConfig As String, sCodigo As String, lNumInt As Long, Optional iQtde As Integer = 1) As Long
'obter um numero automatico a partir de um contador identificado por uma chave
'iQtde identifica quantos numeros devem ser "liberados" a partir do lNumInt retornado
'TEM QUE SER CHAMADA DENTRO DE TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Config_ObterNumInt
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CONFIG_OBTERNUMINT", X)
    If lErro <> SUCESSO Then Error 19993
    
    
    
    If IsMissing(iQtde) Then
        Config_ObterNumInt = X.Config_ObterNumInt(sTabelaConfig, sCodigo, lNumInt)
    Else
        Config_ObterNumInt = X.Config_ObterNumInt(sTabelaConfig, sCodigo, lNumInt, iQtde)
    End If
    
    Exit Function
    
Erro_Config_ObterNumInt:

    Config_ObterNumInt = Err
    
    Select Case Err

        Case 19993
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142108)
        
    End Select
    
    Exit Function

End Function

Function Cod_Nomes_Le(sTabela As String, sCampo_Codigo As String, sCampo_Nome As String, ByVal iTamanho_Nome As Integer, colCodigoDescricao As AdmColCodigoNome) As Long
'Le todos os campos sCampo_Codigo e sCampo_Nome (de tamanho = iTamanho_Nome) da tabela sTabela e coloca na colecao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cod_Nomes_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_COD_NOMES_LE", X)
    If lErro <> SUCESSO Then Error 6646
    
    
    Cod_Nomes_Le = X.Cod_Nomes_Le(sTabela, sCampo_Codigo, sCampo_Nome, iTamanho_Nome, colCodigoDescricao)
    
    Exit Function
    
Erro_Cod_Nomes_Le:

    Cod_Nomes_Le = Err
    
    Select Case Err

        Case 6646
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142109)
        
    End Select
    
    Exit Function

End Function

Function LCod_Nomes_Le(sTabela As String, sCampo_Codigo As String, sCampo_Nome As String, ByVal iTamanho_Nome As Integer, collCodigoDescricao As AdmCollCodigoNome) As Long
'Le todos os campos sCampo_Codigo e sCampo_Nome (de tamanho = iTamanho_Nome) da tabela sTabela e coloca na colecao
'O codigo tem que ser do tipo long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LCod_Nomes_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_LCOD_NOMES_LE", X)
    If lErro <> SUCESSO Then Error 14337
    
    
    LCod_Nomes_Le = X.LCod_Nomes_Le(sTabela, sCampo_Codigo, sCampo_Nome, iTamanho_Nome, collCodigoDescricao)
    
    Exit Function
    
Erro_LCod_Nomes_Le:

    LCod_Nomes_Le = Err
    
    Select Case Err

        Case 14337
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142110)
        
    End Select
    
    Exit Function

End Function

Function GrupoBrowseCampo_Le(sCodGrupo As String, sNomeTela As String, colGrupoBrowseCampo As Collection) As Long
'le os campos selecionados para o grupo x tela em questão em coloca-os na coleção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GrupoBrowseCampo_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_GRUPOBROWSECAMPO_LE", X)
    If lErro <> SUCESSO Then Error 9081
    
    
    GrupoBrowseCampo_Le = X.GrupoBrowseCampo_Le(sCodGrupo, sNomeTela, colGrupoBrowseCampo)
    
    Exit Function
    
Erro_GrupoBrowseCampo_Le:

    GrupoBrowseCampo_Le = Err
    
    Select Case Err

        Case 9081
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142111)
        
    End Select
    
    Exit Function

End Function

Function BrowseIndiceSegmentos_Le(sNomeTela As String, iIndice As Integer, colBrowseIndiceSegmentos As Collection, colBrowseIndice As Collection) As Long
'le todos os segmentos do indice especificado e coloca os resultados na coleção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BrowseIndiceSegmentos_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSEINDICESEGMENTOS_LE", X)
    If lErro <> SUCESSO Then Error 9146
    
    
    BrowseIndiceSegmentos_Le = X.BrowseIndiceSegmentos_Le(sNomeTela, iIndice, colBrowseIndiceSegmentos, colBrowseIndice)
    
    Exit Function
    
Erro_BrowseIndiceSegmentos_Le:

    BrowseIndiceSegmentos_Le = Err
    
    Select Case Err

        Case 9146
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142112)
        
    End Select
    
    Exit Function

End Function

Function Browse_Le_Posicao(lPosicao As Long, sSQL As String, colBrowseValorCampo As Collection, colSelecao As Collection, objBrowse As AdmBrowse) As Long
'le o registro referente a posicao passada como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Browse_Le_Posicao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSE_LE_POSICAO", X)
    If lErro <> SUCESSO Then Error 9164
    
    
    Browse_Le_Posicao = X.Browse_Le_Posicao(lPosicao, sSQL, colBrowseValorCampo, colSelecao, objBrowse)
    
    Exit Function
    
Erro_Browse_Le_Posicao:

    Browse_Le_Posicao = Err
    
    Select Case Err

        Case 9164
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142113)
        
    End Select
    
    Exit Function

End Function

Function BrowseArquivo_Le(sNomeTela As String, objBrowseArquivo As AdmBrowseArquivo) As Long
'le todos os campos da tela de browse para o usuario especificado e coloca os resultados na coleção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BrowseArquivo_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSEARQUIVO_LE", X)
    If lErro <> SUCESSO Then Error 9255
    
    
    BrowseArquivo_Le = X.BrowseArquivo_Le(sNomeTela, objBrowseArquivo)
    
    Exit Function
    
Erro_BrowseArquivo_Le:

    BrowseArquivo_Le = Err
    
    Select Case Err

        Case 9255
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142114)
        
    End Select
    
    Exit Function

End Function

Function Rel_Executar_RotinaAuxiliar(sRotinaAuxiliar As String, objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
'executa rotina preparatoria para a execucao de relatorio, tipicamente abre dialogo para a obtencao de parametros

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Rel_Executar_RotinaAuxiliar
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao(sRotinaAuxiliar, X)
    If lErro <> SUCESSO Then Error 7059
    
    
    
    Select Case sRotinaAuxiliar
    
        Case "CTB_REL_BALVERIF"
            Rel_Executar_RotinaAuxiliar = X.Rel_BalVerif(objRelatorio, objRelOpcoes)
        Case "CTB_REL_BALPATR"
            Rel_Executar_RotinaAuxiliar = X.Rel_BalPatr(objRelatorio, objRelOpcoes)
        Case "CTB_REL_BALPATRPER"
            Rel_Executar_RotinaAuxiliar = X.Rel_BalPatrPer(objRelatorio, objRelOpcoes)
        Case "CTB_REL_DEMRESEX"
            Rel_Executar_RotinaAuxiliar = X.Rel_DemResEx(objRelatorio, objRelOpcoes)
        Case "CTB_REL_DEMRESPER"
            Rel_Executar_RotinaAuxiliar = X.Rel_DemResPer(objRelatorio, objRelOpcoes)
        Case "CTB_REL_DESPCCL"
            Rel_Executar_RotinaAuxiliar = X.Rel_DespPerCcl(objRelatorio, objRelOpcoes)
        Case "CTB_REL_DIARIO"
            Rel_Executar_RotinaAuxiliar = X.Rel_Diario(objRelatorio, objRelOpcoes)
        Case "CTB_REL_LANCCCL"
            Rel_Executar_RotinaAuxiliar = X.Rel_LancCcl(objRelatorio, objRelOpcoes)
        Case "CTB_REL_LANCDATA"
            Rel_Executar_RotinaAuxiliar = X.Rel_LancData(objRelatorio, objRelOpcoes)
        Case "CTB_REL_LANPENDENTE"
            Rel_Executar_RotinaAuxiliar = X.Rel_LanPend(objRelatorio, objRelOpcoes)
        Case "CTB_REL_LANCLOTE"
            Rel_Executar_RotinaAuxiliar = X.Rel_LancLote(objRelatorio, objRelOpcoes)
        Case "CTB_REL_LOTES"
            Rel_Executar_RotinaAuxiliar = X.Rel_Lote(objRelatorio, objRelOpcoes)
        Case "CTB_REL_LOTES_PENDENTES"
            Rel_Executar_RotinaAuxiliar = X.Rel_LotePend(objRelatorio, objRelOpcoes)
        Case "CTB_REL_ORCREAL"
            Rel_Executar_RotinaAuxiliar = X.Rel_OrcReal(objRelatorio, objRelOpcoes)
        Case "CTB_REL_ORCREAL_CCL"
            Rel_Executar_RotinaAuxiliar = X.Rel_OrcRealCcl(objRelatorio, objRelOpcoes)
        Case "CTB_REL_PLANO"
            Rel_Executar_RotinaAuxiliar = X.Rel_PlanoContas(objRelatorio, objRelOpcoes)
        Case "CTB_REL_RAZAO"
            Rel_Executar_RotinaAuxiliar = X.Rel_Razao(objRelatorio, objRelOpcoes)
        Case "CTB_REL_PLANSALDOS"
            Rel_Executar_RotinaAuxiliar = X.Rel_PlanSaldos(objRelatorio, objRelOpcoes)
        Case "CTB_REL_MUTACAOPL"
            Rel_Executar_RotinaAuxiliar = X.Rel_MutacaoPL(objRelatorio, objRelOpcoes)
        Case "CTB_REL_BALPATRCOM"
            Rel_Executar_RotinaAuxiliar = X.Rel_BalPatrCom(objRelatorio, objRelOpcoes)
        Case "CP_REL_DIARIO"
            Rel_Executar_RotinaAuxiliar = X.Rel_CPDiario(objRelatorio, objRelOpcoes)
        Case "CP_REL_RAZAO"
            Rel_Executar_RotinaAuxiliar = X.Rel_CPRazao(objRelatorio, objRelOpcoes)
        Case "CP_REL_CADFORN"
            Rel_Executar_RotinaAuxiliar = X.Rel_CadForn(objRelatorio, objRelOpcoes)
        Case "CP_REL_TITPAG"
            Rel_Executar_RotinaAuxiliar = X.Rel_TitPag(objRelatorio, objRelOpcoes)
        Case "CP_REL_BAIXASCP"
            Rel_Executar_RotinaAuxiliar = X.Rel_BaixasCP(objRelatorio, objRelOpcoes)
        Case "CP_REL_PAGTOSCANCELADOS"
            Rel_Executar_RotinaAuxiliar = X.Rel_PagtosCancelados(objRelatorio, objRelOpcoes)
        Case "CP_REL_RELCHEQUES"
            Rel_Executar_RotinaAuxiliar = X.Rel_RelCheques(objRelatorio, objRelOpcoes)
        Case "CP_REL_POSFORN"
            Rel_Executar_RotinaAuxiliar = X.Rel_PosForn(objRelatorio, objRelOpcoes)
        Case "CP_REL_NFPAGSEMFAT"
            Rel_Executar_RotinaAuxiliar = X.Rel_NFPagSemFat(objRelatorio, objRelOpcoes)
        Case "CP_REL_BORDEROPAG"
            Rel_Executar_RotinaAuxiliar = X.Rel_BorderoPag(objRelatorio, objRelOpcoes)
        Case "CP_REL_CLIATRASO"
            Rel_Executar_RotinaAuxiliar = X.Rel_CliAtraso(objRelatorio, objRelOpcoes)
        Case "CP_REL_DIRF"
            Rel_Executar_RotinaAuxiliar = X.Rel_DIRF(objRelatorio, objRelOpcoes)
        Case "CP_REL_ICMS"
            Rel_Executar_RotinaAuxiliar = X.Rel_ICMS(objRelatorio, objRelOpcoes)
        Case "CP_REL_INSSRet"
            Rel_Executar_RotinaAuxiliar = X.Rel_INSSRet(objRelatorio, objRelOpcoes)
        Case "CP_REL_IRRF"
            Rel_Executar_RotinaAuxiliar = X.Rel_IRRF(objRelatorio, objRelOpcoes)
        Case "CR_REL_COBMALA"
            Rel_Executar_RotinaAuxiliar = X.Rel_Cobrmala(objRelatorio, objRelOpcoes)
        Case "CR_REL_COBTEL"
            Rel_Executar_RotinaAuxiliar = X.Rel_Cobrtel(objRelatorio, objRelOpcoes)
        Case "CR_REL_COMIS_VEND"
            Rel_Executar_RotinaAuxiliar = X.Rel_ComisVen(objRelatorio, objRelOpcoes)
        Case "CR_REL_DEVEDORES"
            Rel_Executar_RotinaAuxiliar = X.Rel_Devedor(objRelatorio, objRelOpcoes)
        Case "CR_REL_DIARIO"
            Rel_Executar_RotinaAuxiliar = X.Rel_CRDiario(objRelatorio, objRelOpcoes)
        Case "CR_REL_POSCLI"
            Rel_Executar_RotinaAuxiliar = X.Rel_PosCli(objRelatorio, objRelOpcoes)
        Case "CR_REL_RAZAO"
            Rel_Executar_RotinaAuxiliar = X.Rel_CRRazao(objRelatorio, objRelOpcoes)
        Case "CR_REL_RESTITREC"
            Rel_Executar_RotinaAuxiliar = X.Rel_ResTitCr(objRelatorio, objRelOpcoes)
        Case "CR_REL_RESUMO_COMIS"
            Rel_Executar_RotinaAuxiliar = X.Rel_ResComis(objRelatorio, objRelOpcoes)
        Case "CR_REL_TITREC"
            Rel_Executar_RotinaAuxiliar = X.Rel_TitRec(objRelatorio, objRelOpcoes)
        Case "CR_REL_CADCLI"
            Rel_Executar_RotinaAuxiliar = X.Rel_CadCli(objRelatorio, objRelOpcoes)
        Case "CR_REL_CADVEND"
            Rel_Executar_RotinaAuxiliar = X.Rel_CadVend(objRelatorio, objRelOpcoes)
        Case "CR_REL_JUROSREC"
            Rel_Executar_RotinaAuxiliar = X.Rel_JurosRec(objRelatorio, objRelOpcoes)
        Case "TES_REL_HISTAPLIC"
            Rel_Executar_RotinaAuxiliar = X.Rel_HistApl(objRelatorio, objRelOpcoes)
        Case "TES_REL_MOVFIN"
            Rel_Executar_RotinaAuxiliar = X.Rel_MovFin(objRelatorio, objRelOpcoes)
        Case "TES_REL_MOVFINDET"
            Rel_Executar_RotinaAuxiliar = X.Rel_MovFinDt(objRelatorio, objRelOpcoes)
        Case "TES_REL_POSAPLIC"
            Rel_Executar_RotinaAuxiliar = X.Rel_PosAplic(objRelatorio, objRelOpcoes)
        Case "EST_REL_PRODUTOS"
            Rel_Executar_RotinaAuxiliar = X.Rel_Produtos(objRelatorio, objRelOpcoes)
        Case "EST_REL_SALDOESTOQUE"
            Rel_Executar_RotinaAuxiliar = X.Rel_SaldoEstoque(objRelatorio, objRelOpcoes)
        Case "EST_REL_ANALISEMOVEST"
            Rel_Executar_RotinaAuxiliar = X.Rel_AnaliseMovEst(objRelatorio, objRelOpcoes)
       Case "EST_REL_BOLETIMENTRADA"
            Rel_Executar_RotinaAuxiliar = X.Rel_BoletimEntrada(objRelatorio, objRelOpcoes)
       Case "EST_REL_CUSTOREP"
            Rel_Executar_RotinaAuxiliar = X.Rel_CustoRep(objRelatorio, objRelOpcoes)
       Case "EST_REL_ETIQINV"
            Rel_Executar_RotinaAuxiliar = X.Rel_EtiqInv(objRelatorio, objRelOpcoes)
       Case "EST_REL_FALTAS"
            Rel_Executar_RotinaAuxiliar = X.Rel_Faltas(objRelatorio, objRelOpcoes)
       Case "EST_REL_KARDEX"
            Rel_Executar_RotinaAuxiliar = X.Rel_Kardex(objRelatorio, objRelOpcoes)
       Case "EST_REL_KARDEXDIA"
            Rel_Executar_RotinaAuxiliar = X.Rel_KardexDia(objRelatorio, objRelOpcoes)
       Case "EST_REL_LISTINV"
            Rel_Executar_RotinaAuxiliar = X.Rel_ListInv(objRelatorio, objRelOpcoes)
       Case "EST_REL_MOVINT"
            Rel_Executar_RotinaAuxiliar = X.Rel_MovInt(objRelatorio, objRelOpcoes)
       Case "EST_REL_ORDEMPRODUCAO"
            Rel_Executar_RotinaAuxiliar = X.Rel_OrdemProducao(objRelatorio, objRelOpcoes)
       Case "EST_REL_PICKLIST"
            Rel_Executar_RotinaAuxiliar = X.Rel_PickList(objRelatorio, objRelOpcoes)
       Case "FAT_REL_PRECOS"
            Rel_Executar_RotinaAuxiliar = X.Rel_Precos(objRelatorio, objRelOpcoes)
       Case "EST_REL_PRODVEND"
            Rel_Executar_RotinaAuxiliar = X.Rel_ProdVend(objRelatorio, objRelOpcoes)
       Case "EST_REL_REQCONSUMO"
            Rel_Executar_RotinaAuxiliar = X.Rel_ReqConsumo(objRelatorio, objRelOpcoes)
       Case "EST_REL_RESESVALOR"
            Rel_Executar_RotinaAuxiliar = X.Rel_ResESValor(objRelatorio, objRelOpcoes)
       Case "EST_REL_VENDASMES"
            Rel_Executar_RotinaAuxiliar = X.Rel_VendasMes(objRelatorio, objRelOpcoes)
       Case "EST_REL_MOVSESTOP"
            Rel_Executar_RotinaAuxiliar = X.Rel_MovsEstOP(objRelatorio, objRelOpcoes)
       Case "EST_REL_REALSTD"
            Rel_Executar_RotinaAuxiliar = X.Rel_RealStd(objRelatorio, objRelOpcoes)
       Case "EST_REL_PONTOPEDIDO"
            Rel_Executar_RotinaAuxiliar = X.Rel_PontoPedido(objRelatorio, objRelOpcoes)
       Case "EST_REL_ANALISEEST"
            Rel_Executar_RotinaAuxiliar = X.Rel_AnaliseEst(objRelatorio, objRelOpcoes)
       Case "EST_REL_KITS"
            Rel_Executar_RotinaAuxiliar = X.Rel_Kit(objRelatorio, objRelOpcoes)
       Case "EST_REL_USOPROD"
            Rel_Executar_RotinaAuxiliar = X.Rel_UsoProd(objRelatorio, objRelOpcoes)
       Case "EST_REL_RESUMO_KARDEX"
            Rel_Executar_RotinaAuxiliar = X.Rel_ResumoKardex(objRelatorio, objRelOpcoes)
       Case "EST_REL_CLASSIFABC"
            Rel_Executar_RotinaAuxiliar = X.Rel_ClassifABC(objRelatorio, objRelOpcoes)
       Case "EST_REL_REGINVMOD7"
            Rel_Executar_RotinaAuxiliar = X.Rel_RegInvMod7(objRelatorio, objRelOpcoes)
       Case "EST_REL_DEMAPINV"
            Rel_Executar_RotinaAuxiliar = X.Rel_DemApInv(objRelatorio, objRelOpcoes)
       Case "FAT_REL_CONTASCORRENTES"
            Rel_Executar_RotinaAuxiliar = X.Rel_ContasCorrentes(objRelatorio, objRelOpcoes)
       Case "FAT_REL_DESPACHO"
            Rel_Executar_RotinaAuxiliar = X.Rel_Despacho(objRelatorio, objRelOpcoes)
       Case "FAT_REL_EMISSAODUPL"
            Rel_Executar_RotinaAuxiliar = X.Rel_EmissaoDupl(objRelatorio, objRelOpcoes)
       Case "FAT_REL_EMISSAOFATURA"
            Rel_Executar_RotinaAuxiliar = X.Rel_EmissaoFatura(objRelatorio, objRelOpcoes)
       Case "FAT_REL_EMISSAONFISCAL"
            Rel_Executar_RotinaAuxiliar = X.Rel_EmissaoNFiscal(objRelatorio, objRelOpcoes)
       Case "FAT_REL_ESTOQUEVENDAS"
            Rel_Executar_RotinaAuxiliar = X.Rel_EstoqueVendas(objRelatorio, objRelOpcoes)
       Case "FAT_REL_FATCLIENTE"
            Rel_Executar_RotinaAuxiliar = X.Rel_FatCliente(objRelatorio, objRelOpcoes)
       Case "FAT_REL_FATCLIENTEPROD"
            Rel_Executar_RotinaAuxiliar = X.Rel_FatClienteProd(objRelatorio, objRelOpcoes)
       Case "FAT_REL_FATCOMISSOES"
            Rel_Executar_RotinaAuxiliar = X.Rel_FatComissoes(objRelatorio, objRelOpcoes)
       Case "FAT_REL_FATVENDEDOR"
            Rel_Executar_RotinaAuxiliar = X.Rel_FatVendedor(objRelatorio, objRelOpcoes)
       Case "FAT_REL_NFISCALDEV"
            Rel_Executar_RotinaAuxiliar = X.Rel_NFiscalDev(objRelatorio, objRelOpcoes)
       Case "FAT_REL_NFISCALTRANSP"
            Rel_Executar_RotinaAuxiliar = X.Rel_NFiscalTransp(objRelatorio, objRelOpcoes)
       Case "FAT_REL_NOTASFISCAIS"
            Rel_Executar_RotinaAuxiliar = X.Rel_NotasFiscais(objRelatorio, objRelOpcoes)
       Case "FAT_REL_PEDAPTOSFAT"
            Rel_Executar_RotinaAuxiliar = X.Rel_PedAptosFat(objRelatorio, objRelOpcoes)
       Case "FAT_REL_PEDNAOENTREGUE"
            Rel_Executar_RotinaAuxiliar = X.Rel_PedNaoEntregue(objRelatorio, objRelOpcoes)
       Case "FAT_REL_PEDPRODUTO"
            Rel_Executar_RotinaAuxiliar = X.Rel_PedProduto(objRelatorio, objRelOpcoes)
       Case "FAT_REL_PEDVENDEDORCLI"
            Rel_Executar_RotinaAuxiliar = X.Rel_PedVendedorCli(objRelatorio, objRelOpcoes)
       Case "FAT_REL_PEDVENDEDORPROD"
            Rel_Executar_RotinaAuxiliar = X.Rel_PedVendedorProd(objRelatorio, objRelOpcoes)
       Case "FAT_REL_PLANREAJUSTE"
            Rel_Executar_RotinaAuxiliar = X.Rel_PlanReajuste(objRelatorio, objRelOpcoes)
       Case "FAT_REL_PRAZOPAGTO"
            Rel_Executar_RotinaAuxiliar = X.Rel_PrazoPagto(objRelatorio, objRelOpcoes)
       Case "FAT_REL_PRENOTA"
            Rel_Executar_RotinaAuxiliar = X.Rel_PreNota(objRelatorio, objRelOpcoes)
       Case "FAT_REL_REALPREV"
            Rel_Executar_RotinaAuxiliar = X.Rel_RealPrev(objRelatorio, objRelOpcoes)
       Case "FAT_REL_RESVENDAS"
            Rel_Executar_RotinaAuxiliar = X.Rel_ResVendas(objRelatorio, objRelOpcoes)
       Case "FAT_REL_TRANSPORTADORAS"
            Rel_Executar_RotinaAuxiliar = X.Rel_Transportadoras(objRelatorio, objRelOpcoes)
       Case "FAT_REL_VENDEDORES"
            Rel_Executar_RotinaAuxiliar = X.Rel_Vendedores(objRelatorio, objRelOpcoes)
       Case "FAT_REL_VISITAS"
            Rel_Executar_RotinaAuxiliar = X.Rel_Visitas(objRelatorio, objRelOpcoes)
        
        Case Else
            Error 19418

    End Select
       
    Rel_Executar_RotinaAuxiliar = SUCESSO
    
    Exit Function
    
Erro_Rel_Executar_RotinaAuxiliar:

    Rel_Executar_RotinaAuxiliar = Err
    
    Select Case Err

        Case 19418
            
        
        Case 7059
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142115)
        
    End Select
    
    Exit Function

End Function

Function Relatorio_Le(objRelatorio As AdmRelatorio) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Relatorio_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_RELATORIO_LE", X)
    If lErro <> SUCESSO Then Error 7059
    
    
    Relatorio_Le = X.Relatorio_Le(objRelatorio)
           
    Exit Function
    
Erro_Relatorio_Le:

    Relatorio_Le = Err
    
    Select Case Err

        Case 7059
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142116)
        
    End Select
    
    Exit Function

End Function

Function Relatorios_Le_GrupoModulo(colRelatorio As Collection, sGrupo As String, sModulo As String) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Relatorios_Le_GrupoModulo
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_RELATORIOS_LE_GRUPOMODULO", X)
    If lErro <> SUCESSO Then Error 7059
    
    
    Relatorios_Le_GrupoModulo = X.Relatorios_Le_GrupoModulo(colRelatorio, sGrupo, sModulo)
           
    Exit Function
    
Erro_Relatorios_Le_GrupoModulo:

    Relatorios_Le_GrupoModulo = Err
    
    Select Case Err

        Case 7059
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142117)
        
    End Select
    
    Exit Function

End Function

Function Relatorio_Grava(objRelatorio As AdmRelatorio, iOperacao As Integer, sModulo As String) As Long
'grava o Relatorio no banco de dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Relatorio_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_RELATORIO_GRAVA", X)
    If lErro <> SUCESSO Then Error 14082
    
    
    Relatorio_Grava = X.Relatorio_Grava(objRelatorio, iOperacao, sModulo)
           
    Exit Function
    
Erro_Relatorio_Grava:

    Relatorio_Grava = Err
    
    Select Case Err

        Case 14082
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142118)
        
    End Select
    
    Exit Function

End Function

Function Relatorio_Exclui(objRelatorio As AdmRelatorio, sModulo As String) As Long
'Deleta o Relatorio no banco de dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Relatorio_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_RELATORIO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 14083
    
    
    Relatorio_Exclui = X.Relatorio_Exclui(objRelatorio, sModulo)
           
    Exit Function
    
Erro_Relatorio_Exclui:

    Relatorio_Exclui = Err
    
    Select Case Err

        Case 14083
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142119)
        
    End Select
    
    Exit Function

End Function

Function RelOpcoes_Le(objRelOpcoes As AdmRelOpcoes) As Long
'le uma opção de relatório contida na tabela RelatorioOpcoes
'objRelOpcoes.sCodRel e objRelOpcoes.sNome tem que vir preenchidos

Dim X As Object, lErro As Long

On Error GoTo Erro_RelOpcoes_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CTB_RELOPCOES_LE", X)
    If lErro <> SUCESSO Then Error 14076
    
    
    RelOpcoes_Le = X.RelOpcoes_Le(objRelOpcoes)
    
    Exit Function
    
Erro_RelOpcoes_Le:

    RelOpcoes_Le = Err
    
    Select Case Err

        Case 14076
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142120)
        
    End Select
    
    Exit Function
    
End Function

Function RelOpcoes_Le_Todos(sCodRel As String, colRelParametros As Collection)
'le todos as opções de relatório contidas na tabela RelatorioOpcoes e coloca-os em colRelParametros

Dim X As Object, lErro As Long

On Error GoTo Erro_RelOpcoes_Le_Todos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CTB_RELOPCOES_LE_TODOS", X)
    If lErro <> SUCESSO Then Error 14077
    
    
    RelOpcoes_Le_Todos = X.RelOpcoes_Le_Todos(sCodRel, colRelParametros)
    
    Exit Function
    
Erro_RelOpcoes_Le_Todos:

    RelOpcoes_Le_Todos = Err
    
    Select Case Err

        Case 14077
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142121)
        
    End Select
    
    Exit Function
    
End Function

Function RelOpcoes_Grava(objRelOpcoes As AdmRelOpcoes, iResultado As Integer) As Long
'inclui ou modifica um registro na tabela RelatorioOpcoes
'iResultado retorna GRAVACAO ou MODIFICACAO

Dim X As Object, lErro As Long

On Error GoTo Erro_RelatorioOpcoes_Gravar
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CTB_RELATORIOOPCOES_GRAVAR", X)
    If lErro <> SUCESSO Then Error 14079
    
    
    RelOpcoes_Grava = X.RelOpcoes_Grava(objRelOpcoes, iResultado)
    
    Exit Function
    
Erro_RelatorioOpcoes_Gravar:

    RelOpcoes_Grava = Err
    
    Select Case Err

        Case 14079
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142122)
        
    End Select
    
    Exit Function
    
End Function

Function RelOpcoes_Exclui(objRelOpcoes As AdmRelOpcoes) As Long
'exclui da tabela RelatorioOpcoes o registro referente a opção em questão

Dim X As Object, lErro As Long

On Error GoTo Erro_RelOpcoesExclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CTB_RELOPCOESEXCLUI", X)
    If lErro <> SUCESSO Then Error 14080
    
    
    RelOpcoes_Exclui = X.RelOpcoes_Exclui(objRelOpcoes)
    
    Exit Function
    
Erro_RelOpcoesExclui:

    RelOpcoes_Exclui = Err
    
    Select Case Err

        Case 14080
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142123)
        
    End Select
    
    Exit Function
    
End Function

Function Modulos_Le_Todos(colModulo As Collection) As Long
'le todos os modulos contidos na tabela Modulos no dicionario de dados e coloca-os em colModulo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Modulos_Le_Todos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_MODULOS_LE_TODOS", X)
    If lErro <> SUCESSO Then Error 9645
    
    
    Modulos_Le_Todos = X.Modulos_Le_Todos(colModulo)
    
    Exit Function
    
Erro_Modulos_Le_Todos:

    Modulos_Le_Todos = Err
    
    Select Case Err

        Case 9645
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142124)
        
    End Select
    
    Exit Function

End Function

''Function Preenche_Espacos(lCodigo As Long, sEspacos As String) As Long
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Preenche_Espacos
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_PREENCHE_ESPACOS", X)
''    If lErro <> SUCESSO Then Error 14336
''
''
''    Preenche_Espacos = X.Preenche_Espacos(lCodigo, sEspacos)
''
''    Exit Function
''
''Erro_Preenche_Espacos:
''
''    Preenche_Espacos = Err
''
''    Select Case Err
''
''        Case 14336
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142125)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Valor_Critica(sValor As String) As Long
'''critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Valor_Critica
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_VALOR_CRITICA", X)
''    If lErro <> SUCESSO Then Error 5709
''
''
''    Valor_Critica = X.Valor_Critica(sValor)
''
''    Exit Function
''
''Erro_Valor_Critica:
''
''    Valor_Critica = Err
''
''    Select Case Err
''
''        Case 5709
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142126)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Valor_Inteiro_Critica(sValor As String) As Long
'''critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO. Senão retorna código de erro e não mostra mensagem.
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Valor_Inteiro_Critica
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_VALOR_INTEIRO_CRITICA", X)
''    If lErro <> AD_BOOL_TRUE Then Error 5894
''
''
''    Valor_Inteiro_Critica = X.Valor_Inteiro_Critica(sValor)
''
''    Exit Function
''
''Erro_Valor_Inteiro_Critica:
''
''    Valor_Inteiro_Critica = Err
''
''    Select Case Err
''
''        Case 5894
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142127)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Valor_Positivo_Critica(sValor As String) As Long
'''critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Valor_Positivo_Critica
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_VALOR_POSITIVO_CRITICA", X)
''    If lErro <> SUCESSO Then Error 10144
''
''
''    Valor_Positivo_Critica = X.Valor_Positivo_Critica(sValor)
''
''    Exit Function
''
''Erro_Valor_Positivo_Critica:
''
''    Valor_Positivo_Critica = Err
''
''    Select Case Err
''
''        Case 10144
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142128)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Valor_NaoNegativo_Critica(sValor As String) As Long
'''critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Valor_NaoNegativo_Critica
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_VALOR_NAO_NEGATIVO_CRITICA", X)
''    If lErro <> SUCESSO Then Error 6753
''
''
''    Valor_NaoNegativo_Critica = X.Valor_NaoNegativo_Critica(sValor)
''
''    Exit Function
''
''Erro_Valor_NaoNegativo_Critica:
''
''    Valor_NaoNegativo_Critica = Err
''
''    Select Case Err
''
''        Case 6753
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142129)
''
''    End Select
''
''    Exit Function
''
''End Function

Function FiliaisEmpresas_Le_Empresa(lCodEmpresa As Long, colFiliais As Collection) As Long
'le todas as filiais da empresa em questão e coloca-as em colFiliais

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FiliaisEmpresas_Le_Empresa
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_FILIAISEMPRESAS_LE_EMPRESA", X)
    If lErro <> SUCESSO Then Error 10201
    
    
    FiliaisEmpresas_Le_Empresa = X.FiliaisEmpresas_Le_Empresa(lCodEmpresa, colFiliais)
    
    Exit Function
    
Erro_FiliaisEmpresas_Le_Empresa:

    FiliaisEmpresas_Le_Empresa = Err
    
    Select Case Err

        Case 10201
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142130)
        
    End Select
    
    Exit Function
    
End Function

''Function Dias_Mes(iMes As Integer, iAno As Integer) As Integer
'''Retorna o ultima dia do mes/ano passado como parametro
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Dias_Mes
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_DIAS_MES", X)
''    If lErro <> SUCESSO Then Error 20070
''
''
''    Dias_Mes = X.Dias_Mes(iMes, iAno)
''
''    Exit Function
''
''Erro_Dias_Mes:
''
''    Dias_Mes = Err
''
''    Select Case Err
''
''        Case 20070
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142131)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function SeguraFoco(objControle As Object, Frames As Object, iFrameAtual As Integer, Opcao As Object) As Long
'''Segura o foco em objControle quando o Controle está em uma Frame de Frames dentro do TabStrip Opcao
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_SeguraFoco
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_SEGURA_FOCO", X)
''    If lErro <> SUCESSO Then Error 6996
''
''
''    SeguraFoco = X.SeguraFoco(objControle, Frames, iFrameAtual, Opcao)
''
''    Exit Function
''
''Erro_SeguraFoco:
''
''    SeguraFoco = Err
''
''    Select Case Err
''
''        Case 6996
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142132)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function SCodigo_Extrai(sString As String) As String
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_SCodigo_Extrai
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_SCODIGO_EXTRAI", X)
''    If lErro <> SUCESSO Then Error 24251
''
''
''    SCodigo_Extrai = X.SCodigo_Extrai(sString)
''
''    Exit Function
''
''Erro_SCodigo_Extrai:
''
''    SCodigo_Extrai = Err
''
''    Select Case Err
''
''        Case 24251
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142133)
''
''    End Select
''
''    Exit Function
''
''End Function

Function FilialEmpresa_Le(objFilialEmpresa As AdmFiliais, Optional bDiretoNoBD As Boolean = False) As Long
'carrega dados da FilialEmpresa da tabela FiliaisEmpresa e de Enderecos
'Se objFilialEmpresa.iCodFilial = 0 entao pegar dados de giFilialEmpresa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilialEmpresa_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_FILIALEMPRESA_LE", X)
    If lErro <> SUCESSO Then Error 24636
    
    
    FilialEmpresa_Le = X.FilialEmpresa_Le(objFilialEmpresa, bDiretoNoBD)
    
    Exit Function
    
Erro_FilialEmpresa_Le:

    FilialEmpresa_Le = Err
    
    Select Case Err

        Case 24636
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142134)
        
    End Select
    
    Exit Function

End Function

Function Endereco_Automatico(lCodigo As Long) As Long
'retorna o número do proximo Endereco disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Endereco_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_ENDERECO_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 14317
    
    
    Endereco_Automatico = X.Endereco_Automatico(lCodigo)
    
    Exit Function
    
Erro_Endereco_Automatico:

    Endereco_Automatico = Err
    
    Select Case Err

        Case 14317
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142135)
        
    End Select
    
    Exit Function
    
End Function

Function Endereco_Grava(objEndereco As ClassEndereco) As Long
'grava o endereco no banco de dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Endereco_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_ENDERECO_GRAVA", X)
    If lErro <> SUCESSO Then Error 14320
    
    
    Endereco_Grava = X.Endereco_Grava(objEndereco)
    
    Exit Function
    
Erro_Endereco_Grava:

    Endereco_Grava = Err
    
    Select Case Err

        Case 14320
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142136)
        
    End Select
    
    Exit Function
    
End Function

Function Endereco_Exclui(lCodEndereco As Long) As Long
'Exclui o Endereco cujo código foi passado
'Chamada DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Endereco_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_ENDERECO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 14348
    
    
    Endereco_Exclui = X.Endereco_Exclui(lCodEndereco)
    
    Exit Function
    
Erro_Endereco_Exclui:

    Endereco_Exclui = Err
    
    Select Case Err

        Case 14348
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142137)
        
    End Select
    
    Exit Function
    
End Function

Function Endereco_Le(objEndereco As ClassEndereco) As Long
'le o endereco a partir da chave fornecida em objEndereco. Devolve os dados em objEndereco

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Endereco_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CPR_ENDERECO_LE", X)
    If lErro <> SUCESSO Then Error 14325
    
    
    Endereco_Le = X.Endereco_Le(objEndereco)
    
    Exit Function
    
Erro_Endereco_Le:

    Endereco_Le = Err
    
    Select Case Err

        Case 14325
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142138)
        
    End Select
    
    Exit Function
    
End Function

''Function TabStrip_PodeTrocarTab(iFrameAtual As Integer, Opcao As Object, objForm As Object, Optional iTestaCTB As Integer = 1) As Long
'''provoca o lost focus do controle ativo durante o evento click da TabStrip Opcao
''' para fazer com que o tratamento de erro seja feita com o controle ainda visivel
''' se isto gerar algum erro ou aviso a troca de frame é desfeita
'''Em "Opcao" deve ser passada a TabStrip
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_TabStrip_PodeTrocarTab
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_TABSTRIP_TROCA", X)
''    If lErro <> SUCESSO Then Error 22885
''
''
''    TabStrip_PodeTrocarTab = X.TabStrip_PodeTrocarTab(iFrameAtual, Opcao, objForm, iTestaCTB)
''
''    Exit Function
''
''Erro_TabStrip_PodeTrocarTab:
''
''    TabStrip_PodeTrocarTab = Err
''
''    Select Case Err
''
''        Case 22885
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142139)
''
''    End Select
''
''    Exit Function
''
''End Function

'Function Usuario_Le(objUsuario As ClassUsuario) As Long
''le os dados do usuario passado como parametro e devolve os dados em objUsuario
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Usuario_Le
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("ADM_USUARIO_LE", X)
'    If lErro <> SUCESSO Then Error 36348
'
'
'    Usuario_Le = X.Usuario_Le(objUsuario)
'
'    Exit Function
'
'Erro_Usuario_Le:
'
'    Usuario_Le = Err
'
'    Select Case Err
'
'        Case 36348
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142140)
'
'    End Select
'
'    Exit Function
'
'End Function

Function Usuario_Grava(objUsuario As ClassUsuario) As Long
'inclui ou modifica um registro na tabela Usuario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Usuario_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_USUARIO_GRAVA", X)
    If lErro <> SUCESSO Then Error 36361
    
    
    Usuario_Grava = X.Usuario_Grava(objUsuario)
    
    Exit Function
    
Erro_Usuario_Grava:

    Usuario_Grava = Err
    
    Select Case Err

        Case 36361
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142141)
        
    End Select
    
    Exit Function
    
End Function

'Function Usuario_Grava1(objUsuario As ClassUsuario) As Long
''inclui ou modifica um registro na tabela Usuario
''IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_Usuario_Grava1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("ADM_USUARIO_GRAVA1", X)
'    If lErro <> SUCESSO Then Error 39702
'
'
'    Usuario_Grava1 = X.Usuario_Grava1(objUsuario)
'
'    Exit Function
'
'Erro_Usuario_Grava1:
'
'    Usuario_Grava1 = Err
'
'    Select Case Err
'
'        Case 39702
'
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142142)
'
'    End Select
'
'    Exit Function
'
'End Function

''Function Combo_Item_Igual_Remove(objComboBox As Object) As Long
'''Remove ítem na ComboBox referente ao objComboBox.text passado por parametro
'''ATENÇÃO: Esta função ao selecionar ativa o ClickEvent
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Combo_Item_Igual_Remove
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_COMBO_ITEM_IGUAL_REMOVE", X)
''    If lErro <> SUCESSO Then Error 19610
''
''
''    Combo_Item_Igual_Remove = X.Combo_Item_Igual_Remove(objComboBox)
''
''    Exit Function
''
''Erro_Combo_Item_Igual_Remove:
''
''    Combo_Item_Igual_Remove = Err
''
''    Select Case Err
''
''        Case 19610
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142143)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function List_Item_Igual(objListBox As Object, sString As String) As Long
'''Seleciona ítem na ListBox referente a sString passada por parametro
'''ATENÇÃO: Esta função ao selecionar ativa o ClickEvent
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_List_Item_Igual
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_LIST_ITEM_IGUAL", X)
''    If lErro <> SUCESSO Then Error 19611
''
''
''    List_Item_Igual = X.List_Item_Igual(objListBox, sString)
''
''    Exit Function
''
''Erro_List_Item_Igual:
''
''    List_Item_Igual = Err
''
''    Select Case Err
''
''        Case 19611
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142144)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function List_Item_Igual_Remove(objListBox As Object, sString As String) As Long
'''remove ítem na ListBox referente a sString passada por parametro
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_List_Item_Igual_Remove
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_LIST_ITEM_IGUAL_REMOVE", X)
''    If lErro <> SUCESSO Then Error 19612
''
''
''    List_Item_Igual_Remove = X.List_Item_Igual_Remove(objListBox, sString)
''
''    Exit Function
''
''Erro_List_Item_Igual_Remove:
''
''    List_Item_Igual_Remove = Err
''
''    Select Case Err
''
''        Case 19612
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142145)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function StrParaDbl(sTexto As String) As Double
'''retorna sTexto como double
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_StrParaDbl
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_STRPARADBL", X)
''    If lErro <> SUCESSO Then Error 19613
''
''
''    StrParaDbl = X.StrParaDbl(sTexto)
''
''    Exit Function
''
''Erro_StrParaDbl:
''
''    StrParaDbl = Err
''
''    Select Case Err
''
''        Case 19613
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142146)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function StrParaInt(sTexto As String) As Integer
'''retorna sTexto como Integer
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_StrParaInt
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_STRPARAINT", X)
''    If lErro <> SUCESSO Then Error 19614
''
''
''    StrParaInt = X.StrParaInt(sTexto)
''
''    Exit Function
''
''Erro_StrParaInt:
''
''    StrParaInt = Err
''
''    Select Case Err
''
''        Case 19614
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142147)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function MaskedParaDate(objMasked As Object) As Date
'''retorna objMasked.Text como Date
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_MaskedParaDate
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_MASKEDPARADATE", X)
''    If lErro <> SUCESSO Then Error 19615
''
''
''    MaskedParaDate = X.MaskedParaDate(objMasked)
''
''    Exit Function
''
''Erro_MaskedParaDate:
''
''    MaskedParaDate = Err
''
''    Select Case Err
''
''        Case 19615
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142148)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function StrParaMasked(objMasked As Object, sString As String)
'''coloca sString em MaskedEdit verificando promptinclude
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_StrParaMasked
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_STRPARAMASKED", X)
''    If lErro <> SUCESSO Then Error 19616
''
''
''    StrParaMasked = X.StrParaMasked(objMasked, sString)
''
''    Exit Function
''
''Erro_StrParaMasked:
''
''    StrParaMasked = Err
''
''    Select Case Err
''
''        Case 19616
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142149)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Sub MaskEdBox_TrataGotFocus(objMasked As Object, Optional viAlterado As Variant)
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_MaskEdBox_TrataGotFocus
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_MASKEDBOX_TRATAGOTFOCUS", X)
''    If lErro <> SUCESSO Then Error 54929
''
''
''
''    If Not IsMissing(viAlterado) Then
''        Call X.MaskEdBox_TrataGotFocus(objMasked, viAlterado)
''    Else
''        Call X.MaskEdBox_TrataGotFocus(objMasked)
''    End If
''
''    Exit Sub
''
''Erro_MaskEdBox_TrataGotFocus:
''
''    Select Case Err
''
''        Case 54929
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142150)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
''Sub Combo_Seleciona_ItemData(objComboBox As Object, ByVal lItemData As Long)
'''seleciona o item da combo que tem ItemData igual ao passado por parametro
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Combo_Seleciona_ItemData
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_COMBO_SELECIONA_ITEMDATA", X)
''    If lErro <> SUCESSO Then Error 19617
''
''
''    Call X.Combo_Seleciona_ItemData(objComboBox, lItemData)
''
''    Exit Sub
''
''Erro_Combo_Seleciona_ItemData:
''
''    Select Case Err
''
''        Case 19617
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142151)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
''Sub List_Seleciona_ItemData(objListBox As Object, ByVal lItemData As Long)
'''seleciona o item da listbox que tem ItemData igual ao passado por parametro
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_List_Seleciona_ItemData
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_LIST_SELECIONA_ITEMDATA", X)
''    If lErro <> SUCESSO Then Error 19618
''
''
''    Call X.List_Seleciona_ItemData(objListBox, lItemData)
''
''    Exit Sub
''
''Erro_List_Seleciona_ItemData:
''
''    Select Case Err
''
''        Case 19618
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142152)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
''Sub DateParaMasked(objMasked As Object, dtData As Date)
'''preenche MaskedEdit com Date
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_DateParaMasked
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_DATEPARAMASKED", X)
''    If lErro <> SUCESSO Then Error 19743
''
''
''    Call X.DateParaMasked(objMasked, dtData)
''
''    Exit Sub
''
''Erro_DateParaMasked:
''
''    Select Case Err
''
''        Case 19743
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142153)
''
''    End Select
''
''    Exit Sub
''
''End Sub

Function Config_ObterAutomatico_EmTrans(sTabelaConfig As String, sCodigo As String, sTabelaObjeto As String, sCampo As String, lNumAuto As Long) As Long
'CHAMADA EM TRANSACAO
'retorna em lNumAuto o número automático disponivel do próximo objeto
'o objeto é identificado por sCodigo
'sTabelaObjeto é a Tabela do objeto
'sTabelaConfig é a Tabela de configuração do módulo
'sCampo é o campo na Tabela do objeto que armazena o número/código do objeto.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Config_ObterAutomatico_EmTrans
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CONFIG_OBTERAUTOMATICO_EMTRANS", X)
    If lErro <> SUCESSO Then Error 45555
    
    
    Config_ObterAutomatico_EmTrans = X.Config_ObterAutomatico_EmTrans(sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, lNumAuto)
    
    Exit Function
    
Erro_Config_ObterAutomatico_EmTrans:

    Config_ObterAutomatico_EmTrans = Err
    
    Select Case Err

        Case 45555
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142154)
        
    End Select
    
    Exit Function
    
End Function

Function Config_ObterAutomatico(sTabelaConfig As String, sCodigo As String, sTabelaObjeto As String, sCampo As String, lNumAuto As Long) As Long
'retorna em lNumAuto o número automático disponivel do próximo objeto
'o objeto é identificado por sCodigo
'sTabelaObjeto é a Tabela do objeto
'sTabelaConfig é a Tabela de configuração do módulo
'sCampo é o campo na Tabela do objeto que armazena o número/código do objeto.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Config_ObterAutomatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CONFIG_OBTERAUTOMATICO", X)
    If lErro <> SUCESSO Then Error 19698
    
    
    Config_ObterAutomatico = X.Config_ObterAutomatico(sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, lNumAuto)
    
    Exit Function
    
Erro_Config_ObterAutomatico:

    Config_ObterAutomatico = Err
    
    Select Case Err

        Case 19698
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142155)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le(objProduto As ClassProduto) As Long
'Lê os dados do produto, cujo codigo foi passado como parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE", X)
    If lErro <> SUCESSO Then Error 24268
    
    
    Produto_Le = X.Produto_Le(objProduto)

    Exit Function
    
Erro_Produto_Le:
    
    Produto_Le = Err
    
    Select Case Err

        Case 24268
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142156)
        
    End Select
    
    Exit Function

End Function

Function Produto_Le2(lComando As Long, objProduto As ClassProduto) As Long
'Lê os dados do produto, cujo codigo foi passado como parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le2
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE2", X)
    If lErro <> SUCESSO Then Error 20576
    
    
    Produto_Le2 = X.Produto_Le2(lComando, objProduto)

    Exit Function
    
Erro_Produto_Le2:
    
    Produto_Le2 = Err
    
    Select Case Err

        Case 20576
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142157)
        
    End Select
    
    Exit Function

End Function

Function Produto_Critica(ByVal sProduto As String, objProduto As ClassProduto, iProdutoPreenchido As Integer) As Long
'Verifica se o Produto existe é ATIVO e não é GERENCIAL.
'Devolve os dados do Produto em objProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Critica
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_CRITICA", X)
    If lErro <> SUCESSO Then Error 24343
    
    
    Produto_Critica = X.Produto_Critica(sProduto, objProduto, iProdutoPreenchido)

    Exit Function
    
Erro_Produto_Critica:
    
    Produto_Critica = Err
    
    Select Case Err

        Case 24343
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142158)
        
    End Select
    
    Exit Function

End Function

Function Produto_Critica1(lComando As Long, ByVal sProduto As String, objProduto As ClassProduto, iProdutoPreenchido As Integer) As Long
'Verifica se o Produto existe é ATIVO e não é GERENCIAL.
'Devolve os dados do Produto em objProduto
'Devolve PRODUTO_PREENCHIDO ou PRODUTO_VAZIO em iProdutoPreenchido

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Critica1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_CRITICA1", X)
    If lErro <> SUCESSO Then Error 20585
    
    
    Produto_Critica1 = X.Produto_Critica1(lComando, sProduto, objProduto, iProdutoPreenchido)

    Exit Function
    
Erro_Produto_Critica1:
    
    Produto_Critica1 = Err
    
    Select Case Err

        Case 20585
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142159)
        
    End Select
    
    Exit Function

End Function

Function Produto_Formata(sProduto As String, sProdutoFormatado As String, iProdutoPreenchido As Integer) As Long
'Devolve o produto no formato do BD: zeros a esquerda em segmentos numericos e brancos a direita para segmentos alfabeticos
'A String de Entrada vem com os caracteres delimitadores e brancos(se for o caso). O comprimento dela é o comprimento total da máscara.
'iProdutoPreenchido retornará PRODUTO_VAZIO se nenhum segmento do produto for preenchido,
'caso contrário retornará PRODUTO_PREENCHIDO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Formata
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_FORMATA", X)
    If lErro <> SUCESSO Then Error 36843
    
    
    Produto_Formata = X.Produto_Formata(sProduto, sProdutoFormatado, iProdutoPreenchido)

    Exit Function
    
Erro_Produto_Formata:
    
    Produto_Formata = Err
    
    Select Case Err

        Case 36843
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142160)
        
    End Select
    
    Exit Function

End Function

Function Usuarios_Le_Todos(colUsuarios As Collection) As Long
'le todos os usuários contidos na tabela de Usuarios e coloca os dados em colUsuarios

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Usuarios_Le_Todos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_USUARIOS_LE_TODOS", X)
    If lErro <> SUCESSO Then Error 44003
    
    
    Usuarios_Le_Todos = X.Usuarios_Le_Todos(colUsuarios)

    Exit Function
    
Erro_Usuarios_Le_Todos:
    
    Usuarios_Le_Todos = Err
    
    Select Case Err

        Case 44003
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142161)
        
    End Select
    
    Exit Function

End Function

''Function Long_Critica1(sNumero As String) As Long
'''Critica se é tipo Long positivo (entre 1 e 99 999)
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Long_Critica1
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_LONG_CRITICA1", X)
''    If lErro <> SUCESSO Then Error 44031
''
''
''    Long_Critica1 = X.Long_Critica1(sNumero)
''    Exit Function
''
''Erro_Long_Critica1:
''
''    Long_Critica1 = Err
''
''    Select Case Err
''
''        Case 44031
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142162)
''
''    End Select
''
''    Exit Function
''
''End Function

''Function RelOpcoes_Testa_Combo(ComboBox As Object, sNome As String)
'''verifica se sNOme está na combobox. Se não estiver insere. Usado na interface de relatórios
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_RelOpcoes_Testa_Combo
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_RELOPCOES_TESTA_COMBO", X)
''    If lErro <> AD_BOOL_TRUE Then Error 44106
''
''
''    RelOpcoes_Testa_Combo = X.RelOpcoes_Testa_Combo(ComboBox, sNome)
''    Exit Function
''
''Erro_RelOpcoes_Testa_Combo:
''
''    RelOpcoes_Testa_Combo = Err
''
''    Select Case Err
''
''        Case 44106
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142163)
''
''    End Select
''
''    Exit Function
''
''End Function

''Function Limpa_Relatorio(Frm As Object) As Long
'''limpa as telas de parametros de relatórios
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Limpa_Relatorio
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_LIMPA_RELATORIO", X)
''    If lErro <> AD_BOOL_TRUE Then Error 44107
''
''
''    Limpa_Relatorio = X.Limpa_Relatorio(Frm)
''    Exit Function
''
''Erro_Limpa_Relatorio:
''
''    Limpa_Relatorio = Err
''
''    Select Case Err
''
''        Case 44107
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142164)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Valor_Long_Critica(sValor As String) As Long
'''critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO. Senão retorna código de erro e não mostra mensagem.
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Valor_Long_Critica
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_VALOR_LONG_CRITICA", X)
''    If lErro <> AD_BOOL_TRUE Then Error 44121
''
''
''    Valor_Long_Critica = X.Valor_Long_Critica(sValor)
''    Exit Function
''
''Erro_Valor_Long_Critica:
''
''    Valor_Long_Critica = Err
''
''    Select Case Err
''
''        Case 44121
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142165)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Valor_Date_Critica(sValor As String) As Long
'''critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO. Senão retorna código de erro e não mostra mensagem.
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Valor_Date_Critica
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_VALOR_DATE_CRITICA", X)
''    If lErro <> AD_BOOL_TRUE Then Error 44122
''
''
''    Valor_Date_Critica = X.Valor_Date_Critica(sValor)
''    Exit Function
''
''Erro_Valor_Date_Critica:
''
''    Valor_Date_Critica = Err
''
''    Select Case Err
''
''        Case 44122
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142166)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Valor_Double_Critica(sValor As String) As Long
'''critica o valor passado como parametro e se estiver tudo ok retorna SUCESSO. Senão retorna código de erro e não mostra mensagem.
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Valor_Double_Critica
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_VALOR_DOUBLE_CRITICA", X)
''    If lErro <> AD_BOOL_TRUE Then Error 44123
''
''
''    Valor_Double_Critica = X.Valor_Double_Critica(sValor)
''    Exit Function
''
''Erro_Valor_Double_Critica:
''
''    Valor_Double_Critica = Err
''
''    Select Case Err
''
''        Case 44123
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142167)
''
''    End Select
''
''    Exit Function
''
''End Function

Function Carga_Arvore_Produto(colNodes As Nodes) As Long
'move os dados do plano de contas do banco de dados para a arvore colNodes.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Carga_Arvore_Produto
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CARGA_ARVORE_PRODUTO", X)
    If lErro <> SUCESSO Then Error 44141
    
    
    Carga_Arvore_Produto = X.Carga_Arvore_Produto(colNodes)
    Exit Function
    
Erro_Carga_Arvore_Produto:

    Carga_Arvore_Produto = Err
    
    Select Case Err

        Case 44141
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142168)
        
    End Select
    
    Exit Function
    
End Function

Function Carga_Arvore_Produto_Netos(objNodeAvo As Node, colNodes As Nodes) As Long
'le os produtos netos do produto objNodeAvo e coloca-os na arvore colNodes.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Carga_Arvore_Produto_Netos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CARGA_ARVORE_PRODUTO_NETOS", X)
    If lErro <> SUCESSO Then Error 44142
    
    
    Carga_Arvore_Produto_Netos = X.Carga_Arvore_Produto_Netos(objNodeAvo, colNodes)
    Exit Function
    
Erro_Carga_Arvore_Produto_Netos:

    Carga_Arvore_Produto_Netos = Err
    
    Select Case Err

        Case 44142
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142169)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le_Niveis0e1(colProduto As Collection) As Long
'le todos os produtos de nível 0 e 1 e coloca-os em colProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_Niveis0e1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE_NIVEIS0E1", X)
    If lErro <> SUCESSO Then Error 44143
    
    
    Produto_Le_Niveis0e1 = X.Produto_Le_Niveis0e1(colProduto)
    Exit Function
    
Erro_Produto_Le_Niveis0e1:

    Produto_Le_Niveis0e1 = Err
    
    Select Case Err

        Case 44143
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142170)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le_Netos(sProduto As String, colProduto As Collection) As Long
'le todos os netos de sProduto e coloca-os em colProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_Netos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE_NETOS", X)
    If lErro <> SUCESSO Then Error 44144
    
    
    Produto_Le_Netos = X.Produto_Le_Netos(sProduto, colProduto)
    Exit Function
    
Erro_Produto_Le_Netos:

    Produto_Le_Netos = Err
    
    Select Case Err

        Case 44144
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142171)
        
    End Select
    
    Exit Function
    
End Function

''Function StrParaLong(sTexto As String) As Long
'''retorna sTexto como long
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_StrParaLong
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_STRPARALONG", X)
''    If lErro <> SUCESSO Then Error 41589
''
''
''    StrParaLong = X.StrParaLong(sTexto)
''
''    Exit Function
''
''Erro_StrParaLong:
''
''    StrParaLong = Err
''
''    Select Case Err
''
''        Case 41589
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142172)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function StrParaDate(sTexto As String) As Date
'''retorna sTexto como date
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_StrParaDate
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_STRPARADATE", X)
''    If lErro <> SUCESSO Then Error 41590
''
''
''    StrParaDate = X.StrParaDate(sTexto)
''
''    Exit Function
''
''Erro_StrParaDate:
''
''    StrParaDate = Err
''
''    Select Case Err
''
''        Case 41590
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142173)
''
''    End Select
''
''    Exit Function
''
''End Function
''
Function Config_Obter_Inteiro_Automatico(sTabelaConfig As String, sCodigo As String, sTabelaObjeto As String, sCampo As String, iNumAuto As Integer) As Long
'retorna em iNumAuto(inteiro), o número automático disponivel do próximo objeto
'o objeto é identificado por sCodigo
'sTabelaObjeto é a Tabela do objeto
'sTabelaConfig é a Tabela de configuração do módulo
'sCampo é o campo na Tabela do objeto que armazena o número/código do objeto.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Config_Obter_Inteiro_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CONFIG_OBTER_INTEIRO_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 45209
    
    
    Config_Obter_Inteiro_Automatico = X.Config_Obter_Inteiro_Automatico(sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, iNumAuto)
    
    Exit Function
    
Erro_Config_Obter_Inteiro_Automatico:

    Config_Obter_Inteiro_Automatico = Err
    
    Select Case Err

        Case 45209
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142174)
        
    End Select
    
    Exit Function
    
End Function

Function Carga_Arvore_Produto_Reserva(colNodes As Nodes) As Long
'move os dados do plano de contas do banco de dados para a arvore colNodes.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Carga_Arvore_Produto_Reserva
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CARGA_ARVORE_PRODUTO_RESERVA", X)
    If lErro <> SUCESSO Then Error 44156
    
    
    Carga_Arvore_Produto_Reserva = X.Carga_Arvore_Produto_Reserva(colNodes)
    
    Exit Function
    
Erro_Carga_Arvore_Produto_Reserva:

    Carga_Arvore_Produto_Reserva = Err
    
    Select Case Err

        Case 44156
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142175)
        
    End Select
    
    Exit Function
    
End Function

Function Carga_Arvore_Produto_Netos_Reserva(objNodeAvo As Node, colNodes As Nodes) As Long
'le os produtos netos do produto objNodeAvo e coloca-os na arvore colNodes.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Carga_Arvore_Produto_Netos_Reserva
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CARGA_ARVORE_PRODUTO_NETOS_RESERVA", X)
    If lErro <> SUCESSO Then Error 44157
    
    
    Carga_Arvore_Produto_Netos_Reserva = X.Carga_Arvore_Produto_Netos_Reserva(objNodeAvo, colNodes)
    
    Exit Function
    
Erro_Carga_Arvore_Produto_Netos_Reserva:

    Carga_Arvore_Produto_Netos_Reserva = Err
    
    Select Case Err

        Case 44157
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142176)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le_Niveis0e1_Reserva(colProduto As Collection) As Long
'le todos os produtos de nível 0 e 1 e coloca-os em colProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_Niveis0e1_Reserva
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE_NIVEIS0E1_RESERVA", X)
    If lErro <> SUCESSO Then Error 44158
    
    
    Produto_Le_Niveis0e1_Reserva = X.Produto_Le_Niveis0e1_Reserva(colProduto)
    
    Exit Function
    
Erro_Produto_Le_Niveis0e1_Reserva:

    Produto_Le_Niveis0e1_Reserva = Err
    
    Select Case Err

        Case 44158
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142177)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le_Netos_Reserva(sProduto As String, colProduto As Collection) As Long
'le todos os netos de sProduto e coloca-os em colProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_Netos_Reserva
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE_NETOS_RESERVA", X)
    If lErro <> SUCESSO Then Error 44159
    
    
    Produto_Le_Netos_Reserva = X.Produto_Le_Netos_Reserva(sProduto, colProduto)
    
    Exit Function
    
Erro_Produto_Le_Netos_Reserva:

    Produto_Le_Netos_Reserva = Err
    
    Select Case Err

        Case 44159
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142178)
        
    End Select
    
    Exit Function
    
End Function

Function FilialEmpresa_Le1(lComando As Long, objFilialEmpresa As AdmFiliais) As Long
'carrega dados da FilialEmpresa da tabela FiliaisEmpresa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilialEmpresa_Le1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_FILIALEMPRESA_LE1", X)
    If lErro <> SUCESSO Then Error 44213
    
    
    FilialEmpresa_Le1 = X.FilialEmpresa_Le1(lComando, objFilialEmpresa)
    
    Exit Function
    
Erro_FilialEmpresa_Le1:

    FilialEmpresa_Le1 = Err
    
    Select Case Err

        Case 44213
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142179)
        
    End Select
    
    Exit Function
    
End Function

Function UsuarioModulo_Le_UsuarioEmpresa(colUsuarioModulo As Collection, sUsuario As String, lEmpresa As Long, iFilialEmpresa As Integer) As Long
'le todos os modulos validos para o usuario/empresa/filial passados como parametro e coloca-os em colModulo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_UsuarioModulo_Le_UsuarioEmpresa
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_USUARIOMODULO_LE_USUARIOEMPRESA", X)
    If lErro <> SUCESSO Then Error 44419
    
    
    
    UsuarioModulo_Le_UsuarioEmpresa = X.UsuarioModulo_Le_UsuarioEmpresa(colUsuarioModulo, sUsuario, lEmpresa, iFilialEmpresa)
    
    Exit Function
    
Erro_UsuarioModulo_Le_UsuarioEmpresa:

    UsuarioModulo_Le_UsuarioEmpresa = Err
    
    Select Case Err

        Case 44419
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142180)
        
    End Select
    
    Exit Function
    
End Function

Function Paises_Le(objPais As ClassPais) As Long
'Lê o Pais a partir da chave fornecida em (Código) e retorna os dados em objPais

Dim X As Object, lErro As Long

On Error GoTo Erro_Paises_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PAISES_LE", X)
    If lErro <> SUCESSO Then Error 45393
    
    
    Paises_Le = X.Paises_Le(objPais)
    
    Exit Function
    
Erro_Paises_Le:

    Paises_Le = Err
    
    Select Case Err

        Case 45393
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142181)
        
    End Select
    
    Exit Function
    
End Function

Function Paises_Exclui(iCodigo As Integer) As Long
'Exclui o Pais do Banco de Dados

Dim X As Object, lErro As Long

On Error GoTo Erro_Paises_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PAISES_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45394
    
    
    Paises_Exclui = X.Paises_Exclui(iCodigo)
    
    Exit Function
    
Erro_Paises_Exclui:

    Paises_Exclui = Err
    
    Select Case Err

        Case 45394
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142182)
        
    End Select
    
    Exit Function
    
End Function

Function Paises_Grava(objPais As ClassPais) As Long
'Grava o Pais no Banco de Dados

Dim X As Object, lErro As Long

On Error GoTo Erro_Paises_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PAISES_GRAVA", X)
    If lErro <> SUCESSO Then Error 45395
    
    
    Paises_Grava = X.Paises_Grava(objPais)
    
    Exit Function
    
Erro_Paises_Grava:

    Paises_Grava = Err
    
    Select Case Err

        Case 45395
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142183)
        
    End Select
    
    Exit Function
    
End Function

''Function PercentParaDbl(sTexto As String) As Double
'''retorna percentagem sTexto como double
'''sTexto deve ser número finalizado por %
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_PercentParaDbl
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_PERCENTPARADBL", X)
''    If lErro <> SUCESSO Then Error 45520
''
''
''    PercentParaDbl = X.PercentParaDbl(sTexto)
''
''    Exit Function
''
''Erro_PercentParaDbl:
''
''    PercentParaDbl = Err
''
''    Select Case Err
''
''        Case 45520
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142184)
''
''    End Select
''
''    Exit Function
''
''End Function

Function UsuarioModulos_Le(objUsuarioModulo As ClassUsuarioModulo, collCodigoNome As AdmCollCodigoNome) As Long
'Lê os Módulos acessíveis ao usuário, empresa e filial passados como parâmetro

Dim X As Object, lErro As Long

On Error GoTo Erro_UsuarioModulos_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_USUARIOMODULOS_LE", X)
    If lErro <> SUCESSO Then Error 44631
    
    
    UsuarioModulos_Le = X.UsuarioModulos_Le(objUsuarioModulo, collCodigoNome)
    
    Exit Function
    
Erro_UsuarioModulos_Le:

    UsuarioModulos_Le = Err
    
    Select Case Err

        Case 44631
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142185)
        
    End Select
    
    Exit Function
    
End Function

Function UsuarioItensMenu_Le(objUsuItensMenu As ClassUsuarioItensMenu, colUsuarioItensMenu As Collection) As Long
'Lê os itens de menu acessíveis ao usuário, empresa e filial passados como parâmetro

Dim X As Object, lErro As Long

On Error GoTo Erro_UsuarioItensMenu_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_USUARIOITESMENU_LE", X)
    If lErro <> SUCESSO Then Error 44632
    
    
    UsuarioItensMenu_Le = X.UsuarioItensMenu_Le(objUsuItensMenu, colUsuarioItensMenu)
    
    Exit Function
    
Erro_UsuarioItensMenu_Le:

    UsuarioItensMenu_Le = Err
    
    Select Case Err

        Case 44632
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142186)
        
    End Select
    
    Exit Function
    
End Function

Function MenuItens_Le(colMenuItens As Collection) As Long
'Lê os itens de menu

Dim X As Object, lErro As Long

On Error GoTo Erro_MenuItens_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_MENUITENS_LE", X)
    If lErro <> SUCESSO Then Error 44633
    
    
    MenuItens_Le = X.MenuItens_Le(colMenuItens)
    
    Exit Function
    
Erro_MenuItens_Le:

    MenuItens_Le = Err
    
    Select Case Err

        Case 44633
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142187)
        
    End Select
    
    Exit Function
    
End Function

Function Usuarios_Le(objUsuarios As ClassUsuarios) As Long
'le  o usuário contido na tabela de Usuario pelo parametro CodUsuario e coloca no objeto

Dim X As Object, lErro As Long

On Error GoTo Erro_Usuarios_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_USUARIOS_LE", X)
    If lErro <> SUCESSO Then Error 45576
    
    
    Usuarios_Le = X.Usuarios_Le(objUsuarios)
    
    Exit Function
    
Erro_Usuarios_Le:

    Usuarios_Le = Err
    
    Select Case Err

        Case 45576
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142188)
        
    End Select
    
    Exit Function
    
End Function

Function Usuarios_Le_Lock(lComando As Long, objUsuarios As ClassUsuarios) As Long
'le  o usuário contido na tabela de Usuario pelo parametro CodUsuario e coloca no objeto

Dim X As Object, lErro As Long

On Error GoTo Erro_Usuarios_Le_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_USUARIOS_LE_LOCK", X)
    If lErro <> SUCESSO Then Error 45577
    
    
    Usuarios_Le_Lock = X.Usuarios_Le_Lock(lComando, objUsuarios)
    
    Exit Function
    
Erro_Usuarios_Le_Lock:

    Usuarios_Le_Lock = Err
    
    Select Case Err

        Case 45577
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142189)
        
    End Select
    
    Exit Function
    
End Function

Function Feriado_Le(objFeriado As ClassFeriado) As Long
'Lê o Feriado à partir dos dados passados como parâmetro

Dim X As Object, lErro As Long

On Error GoTo Erro_Feriado_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_FERIADO_LE", X)
    If lErro <> SUCESSO Then Error 45839
    
    
    Feriado_Le = X.Feriado_Le(objFeriado)
    
    Exit Function
    
Erro_Feriado_Le:

    Feriado_Le = Err
    
    Select Case Err

        Case 45839
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142190)
        
    End Select
    
    Exit Function
    
End Function

Function Feriado_Grava(objFeriado As ClassFeriado) As Long
'Grava Feriado

Dim X As Object, lErro As Long

On Error GoTo Erro_Feriado_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_FERIADO_GRAVA", X)
    If lErro <> SUCESSO Then Error 45840
    
    
    Feriado_Grava = X.Feriado_Grava(objFeriado)
    
    Exit Function
    
Erro_Feriado_Grava:

    Feriado_Grava = Err
    
    Select Case Err

        Case 45840
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142191)
        
    End Select
    
    Exit Function
    
End Function

Function Feriado_Exclui(objFeriado As ClassFeriado) As Long
'Exclui Feriado

Dim X As Object, lErro As Long

On Error GoTo Erro_Feriado_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_FERIADO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45841
    
    
    Feriado_Exclui = X.Feriado_Exclui(objFeriado)
    
    Exit Function
    
Erro_Feriado_Exclui:

    Feriado_Exclui = Err
    
    Select Case Err

        Case 45841
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142192)
        
    End Select
    
    Exit Function
    
End Function

Function Feriado_Le_Todos(dtDataInicial As Date, dtDataFinal As Date, colFeriados As Collection) As Long
'Lê todos os Feriados contidos na tabela de Feriados

Dim X As Object, lErro As Long

On Error GoTo Erro_Feriado_Le_Todos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_FERIADO_LE_TODOS", X)
    If lErro <> SUCESSO Then Error 45842
    
    
    Feriado_Le_Todos = X.Feriado_Le_Todos(dtDataInicial, dtDataFinal, colFeriados)
    
    Exit Function
    
Erro_Feriado_Le_Todos:

    Feriado_Le_Todos = Err
    
    Select Case Err

        Case 45842
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142193)
        
    End Select
    
    Exit Function
    
End Function

Function Relatorio_Le_Reimpressao(objRelatorio As AdmRelatorio) As Long
'Retorna o nome do Arquivo de Reimpressão sse houver

Dim X As Object, lErro As Long

On Error GoTo Erro_Relatorio_Le_Reimpressao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_RELATORIO_LE_REIMPRESSAO", X)
    If lErro <> SUCESSO Then Error 45843
    
    
    Relatorio_Le_Reimpressao = X.Relatorio_Le_Reimpressao(objRelatorio)
    
    Exit Function
    
Erro_Relatorio_Le_Reimpressao:

    Relatorio_Le_Reimpressao = Err
    
    Select Case Err

        Case 45843
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142194)
        
    End Select
    
    Exit Function
    
End Function

Function ModuloFilEmp_Quer_Config(lCodEmpresa As Long, iFilialEmpresa As Integer, colModulos As Collection) As Long
'devolve os módulos que devem ser configurados para a empresa/filial em questãp

Dim X As Object, lErro As Long

On Error GoTo Erro_ModuloFilEmp_Quer_Config
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_MODULOFILEMP_QUER_CONFIG", X)
    If lErro <> SUCESSO Then Error 44852
    
    
    ModuloFilEmp_Quer_Config = X.ModuloFilEmp_Quer_Config(lCodEmpresa, iFilialEmpresa, colModulos)
    
    Exit Function
    
Erro_ModuloFilEmp_Quer_Config:

    ModuloFilEmp_Quer_Config = Err
    
    Select Case Err

        Case 44852
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142195)
        
    End Select
    
    Exit Function
    
End Function
''
''Function ModuloFilEmp_Le_EmpresaFilial(lCodEmpresa As Long, iFilialEmpresa As Integer, colModuloFilEmp As Collection) As Long
'''le todos os objetos ModuloFilEmp para a empresa/filial em questãp e coloca-os em colModuloFilEmp
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_ModuloFilEmp_Le_EmpresaFilial
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_MODULOFILEMP_LE_EMPRESAFILIAL", X)
''    If lErro <> SUCESSO Then Error 44857
''
''
''    ModuloFilEmp_Le_EmpresaFilial = X.ModuloFilEmp_Le_EmpresaFilial(lCodEmpresa, iFilialEmpresa, colModuloFilEmp)
''
''    Exit Function
''
''Erro_ModuloFilEmp_Le_EmpresaFilial:
''
''    ModuloFilEmp_Le_EmpresaFilial = Err
''
''    Select Case Err
''
''        Case 44857
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142196)
''
''    End Select
''
''    Exit Function
''
''End Function

Function Carga_Arvore_Produto_Componentes(colNodes As Nodes) As Long
'Move os dados de Produtos do BD que podem participar de kit e são ativos para a arvore colNodes.

Dim X As Object, lErro As Long

On Error GoTo Erro_Carga_Arvore_Produto_Componentes
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CARGA_ARVORE_PRODUTO_COMPONENTES", X)
    If lErro <> SUCESSO Then Error 45872
    
    
    Carga_Arvore_Produto_Componentes = X.Carga_Arvore_Produto_Componentes(colNodes)
    
    Exit Function
    
Erro_Carga_Arvore_Produto_Componentes:

    Carga_Arvore_Produto_Componentes = Err
    
    Select Case Err

        Case 45872
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142197)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le_Niveis0e1_Componentes(colProduto As Collection) As Long
'le todos os produtos de nível 0 e 1 e coloca-os em colProduto
'produtos que podem participar de kit

Dim X As Object, lErro As Long

On Error GoTo Erro_Produto_Le_Niveis0e1_Componentes
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE_NIVEIS0E1_COMPONENTES", X)
    If lErro <> SUCESSO Then Error 45873
    
    
    Produto_Le_Niveis0e1_Componentes = X.Produto_Le_Niveis0e1_Componentes(colProduto)
    
    Exit Function
    
Erro_Produto_Le_Niveis0e1_Componentes:

    Produto_Le_Niveis0e1_Componentes = Err
    
    Select Case Err

        Case 45873
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142198)
        
    End Select
    
    Exit Function
    
End Function

Function Carga_Arvore_Produto_Netos_Componentes(objNodeAvo As Node, colNodes As Nodes) As Long
'le os produtos netos do produto objNodeAvo e coloca-os na arvore colNodes.

Dim X As Object, lErro As Long

On Error GoTo Erro_Carga_Arvore_Produto_Netos_Componentes
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CARGA_ARVORE_PROD_NETOS_COMPONENTES", X)
    If lErro <> SUCESSO Then Error 45874
    
    
    Carga_Arvore_Produto_Netos_Componentes = X.Carga_Arvore_Produto_Netos_Componentes(objNodeAvo, colNodes)
    
    Exit Function
    
Erro_Carga_Arvore_Produto_Netos_Componentes:

    Carga_Arvore_Produto_Netos_Componentes = Err
    
    Select Case Err

        Case 45874
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142199)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le_Netos_Componentes(sProduto As String, colProduto As Collection) As Long
'le todos os netos de sProduto e coloca-os em colProduto
'produtos que podem ser Kit(Intermediario ou basico)

Dim X As Object, lErro As Long

On Error GoTo Erro_Produto_Le_Netos_Componentes
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE_NETOS_COMPONENTES", X)
    If lErro <> SUCESSO Then Error 45875
    
    
    Produto_Le_Netos_Componentes = X.Produto_Le_Netos_Componentes(sProduto, colProduto)
    
    Exit Function
    
Erro_Produto_Le_Netos_Componentes:

    Produto_Le_Netos_Componentes = Err
    
    Select Case Err

        Case 45875
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142200)
        
    End Select
    
    Exit Function
    
End Function



Function ModuloFilEmp_Atualiza_Configurado(lCodEmpresa As Long, iFilialEmpresa As Integer, colModuloFilEmp As Collection) As Long
'atualiza a flag configurado de todos os modulos (sigla) passados em colMOduloFilEmp para a Empresa e Filial em questão
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim X As Object, lErro As Long

On Error GoTo Erro_ModuloFilEmp_Atualiza_Configurado
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_MODULOFILEMP_ATUALIZA_CONFIGURADO", X)
    If lErro <> SUCESSO Then Error 44955
    
    
    ModuloFilEmp_Atualiza_Configurado = X.ModuloFilEmp_Atualiza_Configurado(lCodEmpresa, iFilialEmpresa, colModuloFilEmp)
    
    Exit Function
    
Erro_ModuloFilEmp_Atualiza_Configurado:

    ModuloFilEmp_Atualiza_Configurado = Err
    
    Select Case Err

        Case 44955
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142201)
        
    End Select
    
    Exit Function
    
End Function

Public Function Modulos_Le_Empresa_Filial(ByVal lEmpresa As Long, ByVal iFilialEmpresa As Integer, colModulo As AdmColModulo) As Long
'Le todos os modulos associados a lEmpresa/iFilialEmpresa contidos na tabela Modulos no dicionario de dados e coloca-os em colModulo

Dim X As Object, lErro As Long

On Error GoTo Erro_Modulos_Le_Empresa_Filial
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_MODULOS_LE_EMPRESA_FILIAL", X)
    If lErro <> SUCESSO Then Error 44985
    
    
    Modulos_Le_Empresa_Filial = X.Modulos_Le_Empresa_Filial(lEmpresa, iFilialEmpresa, colModulo)
    
    Exit Function
    
Erro_Modulos_Le_Empresa_Filial:

    Modulos_Le_Empresa_Filial = Err
    
    Select Case Err

        Case 44985
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142202)
        
    End Select
    
    Exit Function
    
End Function

'''Seleciona o ítem da ListBox correspondente a sTextBox (MaskedEdit ou TextBox) que pode ser numérico ou string
''Function List_Seleciona(objMaskedEdit As Object, objListBox As Object, iCodigo As Integer) As Long
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_List_Seleciona
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_LIST_SELECIONA", X)
''    If lErro <> SUCESSO Then Error 54502
''
''
''    List_Seleciona = X.List_Seleciona(objMaskedEdit, objListBox, iCodigo)
''
''    Exit Function
''
''Erro_List_Seleciona:
''
''    List_Seleciona = Err
''
''    Select Case Err
''
''        Case 54502
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142203)
''
''    End Select
''
''    Exit Function
''
''End Function
''
'''Seleciona o ítem da Combo correspondente ao Text que pode ser numérico-long ou string
''Function LCombo_Seleciona(objComboBox As Object, lCodigo As Long) As Long
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_LCombo_Seleciona
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_LCOMBO_SELECIONA", X)
''    If lErro <> SUCESSO Then Error 54534
''
''
''    LCombo_Seleciona = X.LCombo_Seleciona(objComboBox, lCodigo)
''
''    Exit Function
''
''Erro_LCombo_Seleciona:
''
''    LCombo_Seleciona = Err
''
''    Select Case Err
''
''        Case 54534
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142204)
''
''    End Select
''
''    Exit Function
''
''End Function

Function Relatorio_AtualizarArqReimp(sCodRel As String, sNomeArqReimp As String) As Long
'atualiza nome do arquivo de reimpressao no cadastro de relatorios do dicionario de dados

Dim X As Object, lErro As Long

On Error GoTo Erro_Relatorio_AtualizarArqReimp
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_RELATORIO_ATUALIZARARQREIMP", X)
    If lErro <> SUCESSO Then Error 54664
    
    
    Relatorio_AtualizarArqReimp = X.Relatorio_AtualizarArqReimp(sCodRel, sNomeArqReimp)
    
    Exit Function
    
Erro_Relatorio_AtualizarArqReimp:

    Relatorio_AtualizarArqReimp = Err
    
    Select Case Err

        Case 54664
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142205)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le_Niveis0e1_Inventariado(colProduto As Collection) As Long
'le todos os produtos de nível 0 e 1 e coloca-os em colProduto

Dim X As Object, lErro As Long

On Error GoTo Erro_Produto_Le_Niveis0e1_Inventariado
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE_NIVEIS0E1_INVENTARIADO", X)
    If lErro <> SUCESSO Then Error 54712
    
    
    Produto_Le_Niveis0e1_Inventariado = X.Produto_Le_Niveis0e1_Inventariado(colProduto)
    
    Exit Function
    
Erro_Produto_Le_Niveis0e1_Inventariado:

    Produto_Le_Niveis0e1_Inventariado = Err
    
    Select Case Err

        Case 54712
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142206)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le_Netos_Inventariado(sProduto As String, colProduto As Collection) As Long
'le todos os netos de sProduto e coloca-os em colProduto

Dim X As Object, lErro As Long

On Error GoTo Erro_Produto_Le_Netos_Inventariado
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE_NETOS_INVENTARIADO", X)
    If lErro <> SUCESSO Then Error 54713
    
    
    Produto_Le_Netos_Inventariado = X.Produto_Le_Netos_Inventariado(sProduto, colProduto)
    
    Exit Function
    
Erro_Produto_Le_Netos_Inventariado:

    Produto_Le_Netos_Inventariado = Err
    
    Select Case Err

        Case 54713
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142207)
        
    End Select
    
    Exit Function
    
End Function

Function Dias_Uteis(dtDataMenor As Date, dtDataMaior As Date, iDias As Integer) As Long
'Retorna o número de dias úteis que existem entre as datas passadas
'???? está longe de ser uma implementacao eficiente. NAO OTIMIZAR AGORA.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Dias_Uteis
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_DIAS_UTEIS", X)
    If lErro <> SUCESSO Then Error 19933
    
    
    Dias_Uteis = X.Dias_Uteis(dtDataMenor, dtDataMaior, iDias)
    
    Exit Function
    
Erro_Dias_Uteis:

    Dias_Uteis = Err
    
    Select Case Err

        Case 19933
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142208)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le_Niveis0e1_Venda(colProduto As Collection) As Long
'le todos os produtos de nível 0 e 1 e coloca-os em colProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_Niveis0e1_Venda
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE_NIVEIS0E1_VENDA", X)
    If lErro <> SUCESSO Then Error 54725
    
    
    Produto_Le_Niveis0e1_Venda = X.Produto_Le_Niveis0e1_Venda(colProduto)
    
    Exit Function
    
Erro_Produto_Le_Niveis0e1_Venda:

    Produto_Le_Niveis0e1_Venda = Err
    
    Select Case Err

        Case 54725
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142209)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le_Netos_Venda(sProduto As String, colProduto As Collection) As Long
'le todos os netos de sProduto e coloca-os em colProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_Netos_Venda
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE_NETOS_VENDA", X)
    If lErro <> SUCESSO Then Error 54726
    
    
    Produto_Le_Netos_Venda = X.Produto_Le_Netos_Venda(sProduto, colProduto)
    
    Exit Function
    
Erro_Produto_Le_Netos_Venda:

    Produto_Le_Netos_Venda = Err
    
    Select Case Err

        Case 54726
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142210)
        
    End Select
    
    Exit Function
    
End Function

Function Carga_Arvore_Produto_Netos_Venda(objNodeAvo As Node, colNodes As Nodes) As Long
'le os produtos netos do produto objNodeAvo e coloca-os na arvore colNodes.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Carga_Arvore_Produto_Netos_Venda
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CARGA_ARVORE_PRODUTO_NETOS_VENDA", X)
    If lErro <> SUCESSO Then Error 54727
    
    
    Carga_Arvore_Produto_Netos_Venda = X.Carga_Arvore_Produto_Netos_Venda(objNodeAvo, colNodes)
    
    Exit Function
    
Erro_Carga_Arvore_Produto_Netos_Venda:

    Carga_Arvore_Produto_Netos_Venda = Err
    
    Select Case Err

        Case 54727
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142211)
        
    End Select
    
    Exit Function
    
End Function

Function Carga_Arvore_Produto_Venda(colNodes As Nodes) As Long
'Move os dados de Produtos do BD que pode ser Vendido e são ativos para a arvore colNodes.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Carga_Arvore_Produto_Venda
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CARGA_ARVORE_PRODUTO_VENDA", X)
    If lErro <> SUCESSO Then Error 54728
    
    
    Carga_Arvore_Produto_Venda = X.Carga_Arvore_Produto_Venda(colNodes)
    
    Exit Function
    
Erro_Carga_Arvore_Produto_Venda:

    Carga_Arvore_Produto_Venda = Err
    
    Select Case Err

        Case 54728
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142212)
        
    End Select
    
    Exit Function
    
End Function

Function Dias_Uteis_Antes(dtDataEntrada As Date, iDias As Integer, dtDataSaida As Date) As Long
'Retorna a data iDias úteis anterior a data de entrada
'???? está longe de ser uma implementacao eficiente. NAO OTIMIZAR AGORA.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Dias_Uteis_Antes
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_DIAS_UTEIS_ANTES", X)
    If lErro <> SUCESSO Then Error 54729
    
    
    Dias_Uteis_Antes = X.Dias_Uteis_Antes(dtDataEntrada, iDias, dtDataSaida)
    
    Exit Function
    
Erro_Dias_Uteis_Antes:

    Dias_Uteis_Antes = Err
    
    Select Case Err

        Case 54729
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142213)
        
    End Select
    
    Exit Function
    
End Function

''Function Porcentagem_Critica_Nao_Zero(sString As String) As Long
'''Critica o valor passado em sString , se nao for entre -100 e 100 ou se for zero retorna erro.
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Porcentagem_Critica_Nao_Zero
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_PORCENTAGEM_CRITICA_NAO_ZERO", X)
''    If lErro <> SUCESSO Then Error 54754
''
''
''    Porcentagem_Critica_Nao_Zero = X.Porcentagem_Critica_Nao_Zero(sString)
''
''    Exit Function
''
''Erro_Porcentagem_Critica_Nao_Zero:
''
''    Porcentagem_Critica_Nao_Zero = Err
''
''    Select Case Err
''
''        Case 54754
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142214)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function PrimeiroDia_Semana(dtData As Date, dtDataPrimeiroDia As Date) As Long
'''Esta funcao retorna o primeiro dia da semana da data em questao
'''"A DATA DA SEGUNDA-FEIRA"
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_PrimeiroDia_Semana
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_PRIMEIRODIA_SEMANA", X)
''    If lErro <> SUCESSO Then Error 54763
''
''
''    PrimeiroDia_Semana = X.PrimeiroDia_Semana(dtData, dtDataPrimeiroDia)
''
''    Exit Function
''
''Erro_PrimeiroDia_Semana:
''
''    PrimeiroDia_Semana = Err
''
''    Select Case Err
''
''        Case 54763
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142215)
''
''    End Select
''
''    Exit Function
''
''End Function

''Function RelOpcoes_ComboOpcoes_Preenche(objRelatorio As AdmRelatorio, ComboOpcoes As Object, objRelOpcoes As AdmRelOpcoes, objTela As Object) As Long
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_RelOpcoes_ComboOpcoes_Preenche
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_RELOPCOES_COMBOOPCOES_PREENCHE", X)
''    If lErro <> SUCESSO Then Error 54764
''
''
''    RelOpcoes_ComboOpcoes_Preenche = X.RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, objTela)
''
''    Exit Function
''
''Erro_RelOpcoes_ComboOpcoes_Preenche:
''
''    RelOpcoes_ComboOpcoes_Preenche = Err
''
''    Select Case Err
''
''        Case 54764
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142216)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Sub RelOpcoes_ComboOpcoes_Click(objRelOpcoes As AdmRelOpcoes, ComboOpcoes As Object, objTela As Object)
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_RelOpcoes_ComboOpcoes_Click
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_RELOPCOES_COMBOOPCOES_CLICK", X)
''    If lErro <> SUCESSO Then Error 54765
''
''
''    Call X.RelOpcoes_ComboOpcoes_Click(objRelOpcoes, ComboOpcoes, objTela)
''
''    Exit Sub
''
''Erro_RelOpcoes_ComboOpcoes_Click:
''
''    Select Case Err
''
''        Case 54765
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142217)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
''Sub RelOpcoes_ComboOpcoes_Validate(ComboOpcoes As Object)
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_RelOpcoes_ComboOpcoes_Validate
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_RELOPCOES_COMBOOPCOES_LOSTFOCUS", X)
''    If lErro <> SUCESSO Then Error 56734
''
''
''    Call X.RelOpcoes_ComboOpcoes_Validate(bSGECancelDummy,ComboOpcoes)
''
''    Exit Sub
''
''Erro_RelOpcoes_ComboOpcoes_Validate:
''
''    Select Case Err
''
''        Case 56734
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142218)
''
''    End Select
''
''    Exit Sub
''
''End Sub

Function Produto_Critica_Filial(ByVal sProduto As String, objProduto As ClassProduto, iProdutoPreenchido As Integer) As Long
'Verifica se o Produto existe é ATIVO e não é GERENCIAL.
'Devolve os dados do Produto em objProduto
'Devolve PRODUTO_PREENCHIDO ou PRODUTO_VAZIO em iProdutoPreenchido
'Verifica em Produto está relacionado com a FilialEmpresa.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Critica_Filial
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_CRITICA_FILIAL", X)
    If lErro <> SUCESSO Then Error 54837
    
    
    Produto_Critica_Filial = X.Produto_Critica_Filial(sProduto, objProduto, iProdutoPreenchido)
    
    Exit Function
    
Erro_Produto_Critica_Filial:

    Produto_Critica_Filial = Err
    
    Select Case Err

        Case 54837
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142219)
        
    End Select
    
    Exit Function
    
End Function

Function Estado_Seleciona(objEstado As Object) As Long
'Seleciona Estado da FilialEmpresa na(s) Combo(s) passada(s)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Estado_Seleciona
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_ESTADO_SELECIONA", X)
    If lErro <> SUCESSO Then Error 54910
    
    
    Estado_Seleciona = X.Estado_Seleciona(objEstado)
    
    Exit Function
    
Erro_Estado_Seleciona:

    Estado_Seleciona = Err
    
    Select Case Err

        Case 54910
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142220)
        
    End Select
    
    Exit Function
    
End Function

''Function Combo_Item_Igual_CI(objComboBox As Object)
'''Seleciona ítem na ComboBox referente ao objComboBox.text passado por parametro
'''Seleciona independente de ser digitada letra maiuscula ou letra minuscula
'''ATENÇÃO: Esta função ao selecionar ativa o ClickEvent
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Combo_Item_Igual_CI
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_COMBO_ITEM_IGUAL_CI", X)
''    If lErro <> SUCESSO Then Error 54930
''
''
''    Combo_Item_Igual_CI = X.Combo_Item_Igual_CI(objComboBox)
''
''    Exit Function
''
''Erro_Combo_Item_Igual_CI:
''
''    Combo_Item_Igual_CI = Err
''
''    Select Case Err
''
''        Case 54930
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142221)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Porcentagem_Critica_Negativa(sString As String) As Long
'''Critica o valor passado em sString aceitando valores negativos até -99.99
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Porcentagem_Critica_Negativa
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("ADM_PORCENTAGEM_CRITICA_NEGATIVA", X)
''    If lErro <> SUCESSO Then Error 54931
''
''
''    Porcentagem_Critica_Negativa = X.Porcentagem_Critica_Negativa(sString)
''
''    Exit Function
''
''Erro_Porcentagem_Critica_Negativa:
''
''    Porcentagem_Critica_Negativa = Err
''
''    Select Case Err
''
''        Case 54931
''
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142222)
''
''    End Select
''
''    Exit Function
''
''End Function

Function Carga_Arvore_Produto_Inventariado(colNodes As Nodes) As Long
'Move os dados de Produtos do BD que têm Produto Inventariado e são ativos para a arvore colNodes.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Carga_Arvore_Produto_Inventariado
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CARGA_ARVORE_PRODUTO_INVENTARIADO", X)
    If lErro <> SUCESSO Then Error 59520
    
    
    Carga_Arvore_Produto_Inventariado = X.Carga_Arvore_Produto_Inventariado(colNodes)
    
    Exit Function
    
Erro_Carga_Arvore_Produto_Inventariado:

    Carga_Arvore_Produto_Inventariado = Err
    
    Select Case Err

        Case 59520
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142223)
        
    End Select
    
    Exit Function
    
End Function

Function Carga_Arvore_Produto_Netos_Inventariado(objNodeAvo As Node, colNodes As Nodes) As Long
'le os produtos netos do produto objNodeAvo e coloca-os na arvore colNodes.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Carga_Arvore_Produto_Netos_Inventariado
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CARGA_ARVORE_PRODUTO_NETOS_INVENT", X)
    If lErro <> SUCESSO Then Error 59521
    
    
    Carga_Arvore_Produto_Netos_Inventariado = X.Carga_Arvore_Produto_Netos_Inventariado(objNodeAvo, colNodes)
    
    Exit Function
    
Erro_Carga_Arvore_Produto_Netos_Inventariado:

    Carga_Arvore_Produto_Netos_Inventariado = Err
    
    Select Case Err

        Case 59521
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142224)
        
    End Select
    
    Exit Function
    
End Function

Function Carga_Arvore_Produto_Netos_Estoque(objNodeAvo As Node, colNodes As Nodes) As Long
'le os produtos netos do produto objNodeAvo e coloca-os na arvore colNodes.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Carga_Arvore_Produto_Netos_Estoque
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CARGA_ARVORE_PRODUTO_NETOS_ESTOQUE", X)
    If lErro <> SUCESSO Then Error 59522
    
    
    Carga_Arvore_Produto_Netos_Estoque = X.Carga_Arvore_Produto_Netos_Estoque(objNodeAvo, colNodes)
    
    Exit Function
    
Erro_Carga_Arvore_Produto_Netos_Estoque:

    Carga_Arvore_Produto_Netos_Estoque = Err
    
    Select Case Err

        Case 59522
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142225)
        
    End Select
    
    Exit Function
    
End Function

Function Carga_Arvore_Produto_Estoque(colNodes As Nodes) As Long
'Move os dados de Produtos do BD que são Estocaveis e são ativos para a arvore colNodes.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Carga_Arvore_Produto_Estoque
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CARGA_ARVORE_PRODUTO_ESTOQUE", X)
    If lErro <> SUCESSO Then Error 59523
    
    
    Carga_Arvore_Produto_Estoque = X.Carga_Arvore_Produto_Estoque(colNodes)
    
    Exit Function
    
Erro_Carga_Arvore_Produto_Estoque:

    Carga_Arvore_Produto_Estoque = Err
    
    Select Case Err

        Case 59523
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142226)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le_Netos_Estoque(sProduto As String, colProduto As Collection) As Long
'le todos os netos de sProduto e coloca-os em colProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_Netos_Estoque
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE_NETOS_ESTOQUE", X)
    If lErro <> SUCESSO Then Error 59524
    
    
    Produto_Le_Netos_Estoque = X.Produto_Le_Netos_Estoque(sProduto, colProduto)
    
    Exit Function
    
Erro_Produto_Le_Netos_Estoque:

    Produto_Le_Netos_Estoque = Err
    
    Select Case Err

        Case 59524
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142227)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le_Niveis0e1_Estoque(colProduto As Collection) As Long
'le todos os produtos de nível 0 e 1 e coloca-os em colProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_Niveis0e1_Estoque
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE_NIVEIS0E1_ESTOQUE", X)
    If lErro <> SUCESSO Then Error 59525
    
    
    Produto_Le_Niveis0e1_Estoque = X.Produto_Le_Niveis0e1_Estoque(colProduto)
    
    Exit Function
    
Erro_Produto_Le_Niveis0e1_Estoque:

    Produto_Le_Niveis0e1_Estoque = Err
    
    Select Case Err

        Case 59525
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142228)
        
    End Select
    
    Exit Function
    
End Function

Function Relatorios_Le_Outros(colRelatorio As Collection, sGrupo As String, sModulo As String) As Long
'retorna os relatorios referentes ao grupo e modulo que o usuario tem acesso e devem aparecer no dialogo de "outros"
'os codigos dos relatorios sao retornados em colRelatorio e o grupo e modulo sao passados como parametros

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Relatorios_Le_Outros
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_RELATORIOS_LE_OUTROS", X)
    If lErro <> SUCESSO Then Error 59529
    
    
    Relatorios_Le_Outros = X.Relatorios_Le_Outros(colRelatorio, sGrupo, sModulo)
    
    Exit Function
    
Erro_Relatorios_Le_Outros:

    Relatorios_Le_Outros = Err
    
    Select Case Err

        Case 59529
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142229)
        
    End Select
    
    Exit Function
    
End Function

Function Campos_Le_Todos(colCampos As Collection) As Long
'le todos os campos e carrega-os na coleção passada como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Campos_Le_Todos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CAMPOS_LE_TODOS", X)
    If lErro <> SUCESSO Then Error 55987
    
    
    Campos_Le_Todos = X.Campos_Le_Todos(colCampos)
    
    Exit Function
    
Erro_Campos_Le_Todos:

    Campos_Le_Todos = Err
    
    Select Case Err

        Case 55987
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142230)
        
    End Select
    
    Exit Function

End Function

Function Retorna_ColFiliais() As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Retorna_ColFiliais
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_RETORNA_COLFILIAIS", X)
    If lErro <> SUCESSO Then Error 60779
    
    
    Retorna_ColFiliais = X.Retorna_ColFiliais()
    
    Exit Function
    
Erro_Retorna_ColFiliais:

    Retorna_ColFiliais = Err
    
    Select Case Err

        Case 60779
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142231)
        
    End Select
    
    Exit Function

End Function

Function BrowseIndiceUsuario_Le(sNomeTela As String, sCodUsuario As String, colBrowseIndice As Collection) As Long
'le todos os indices criados pelo usuario referentes a tela especificada e coloca os resultados na coleção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BrowseIndiceUsuario_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSEINDICEUSUARIO_LE", X)
    If lErro <> SUCESSO Then Error 60826
    
    
    BrowseIndiceUsuario_Le = X.BrowseIndiceUsuario_Le(sNomeTela, sCodUsuario, colBrowseIndice)
    
    Exit Function
    
Erro_BrowseIndiceUsuario_Le:

    BrowseIndiceUsuario_Le = Err
    
    Select Case Err

        Case 60826
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142232)
        
    End Select
    
    Exit Function

End Function

Function BrowseConfiguracoes_Grava(objBrowse As AdmBrowse, objBrowseUsuarioOrdenacao As AdmBrowseUsuarioOrdenacao) As Long
'grava as configurações do browser alteradas pelo usuario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BrowseConfiguracoes_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSECONFIGURACOES_GRAVA", X)
    If lErro <> SUCESSO Then Error 60847
    
    
    BrowseConfiguracoes_Grava = X.BrowseConfiguracoes_Grava(objBrowse, objBrowseUsuarioOrdenacao)
    
    Exit Function
    
Erro_BrowseConfiguracoes_Grava:

    BrowseConfiguracoes_Grava = Err
    
    Select Case Err

        Case 60847
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142233)
        
    End Select
    
    Exit Function

End Function

Function Produto_Le_Netos_Venda_EmpresaToda(sProduto As String, colProduto As Collection) As Long
'le todos os netos de sProduto e coloca-os em colProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_Netos_Venda_EmpresaToda
      
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE_NETOS_VENDA_EMPRESATODA", X)
    If lErro <> SUCESSO Then Error 59856

    Produto_Le_Netos_Venda_EmpresaToda = X.Produto_Le_Netos_Venda_EmpresaToda(sProduto, colProduto)
    
    Exit Function
    
Erro_Produto_Le_Netos_Venda_EmpresaToda:

    Produto_Le_Netos_Venda_EmpresaToda = Err
    
    Select Case Err

        Case 59856
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142234)
        
    End Select
    
    Exit Function

End Function

Function Produto_Le_Niveis0e1_Venda_EmpresaToda(colProduto As Collection) As Long
'Lê todos os produtos que podem ser vendidos para a Empresa Toda e que sejam de nível 0 e 1 e coloca-os em colProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_Niveis0e1_Venda_EmpresaToda
      
    lErro = Rotina_ObterObjetoFuncao("ADM_PRODUTO_LE_NIVEIS0E1_VENDA_EMPRESATO", X)
    If lErro <> SUCESSO Then Error 59857

    Produto_Le_Niveis0e1_Venda_EmpresaToda = X.Produto_Le_Niveis0e1_Venda_EmpresaToda(colProduto)
    
    Exit Function
    
Erro_Produto_Le_Niveis0e1_Venda_EmpresaToda:

    Produto_Le_Niveis0e1_Venda_EmpresaToda = Err
    
    Select Case Err

        Case 59857
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142235)
        
    End Select
    
    Exit Function

End Function

Function EdicaoTela_Le() As Long
'Le as propriedades das Telas no dicionário de Dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EdicaoTela_Le
    
    lErro = Rotina_ObterObjetoFuncao("ADM_EDICAOTELA_LE", X)
    If lErro <> SUCESSO Then Error 64021
    
    EdicaoTela_Le = X.EdicaoTela_Le()
    
    Exit Function
        
Erro_EdicaoTela_Le:

    EdicaoTela_Le = Err
    
    Select Case Err

        Case 64021
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142236)
        
    End Select
    
    Exit Function

End Function

Function EdicaoTela_Grava() As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EdicaoTela_Grava
    
    lErro = Rotina_ObterObjetoFuncao("ADM_EDICAOTELA_GRAVA", X)
    If lErro <> SUCESSO Then Error 65080
    
    EdicaoTela_Grava = X.EdicaoTela_Grava()
    
    Exit Function
        
Erro_EdicaoTela_Grava:

    EdicaoTela_Grava = Err
    
    Select Case Err

        Case 65080
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142237)
        
    End Select
    
    Exit Function

End Function

Function SCombo_Seleciona(objComboBox As Object) As Long
'Seleciona o ítem da Combo correspondente ao Text, procura dos dois lados do Separador
'Retorna o Codigo do item Selecionado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SCombo_Seleciona
    
    lErro = Rotina_ObterObjetoFuncao("ADM_SCOMBO_SELECIONA", X)
    If lErro <> SUCESSO Then Error 65085
    
    SCombo_Seleciona = X.SCombo_Seleciona(objComboBox)
    
    Exit Function
        
Erro_SCombo_Seleciona:

    SCombo_Seleciona = Err
    
    Select Case Err

        Case 65085
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142238)
        
    End Select
    
    Exit Function

End Function

Function Carga_Arvore_Produto_Produziveis(colNodes As Nodes) As Long
'move os dados do plano de contas do banco de dados para a arvore colNodes.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Carga_Arvore_Produto_Produziveis
    
    lErro = Rotina_ObterObjetoFuncao("ADM_CARGA_ARVORE_PRODUTO_PRODUZIVEIS", X)
    If lErro <> SUCESSO Then Error 65100
    
    Carga_Arvore_Produto_Produziveis = X.Carga_Arvore_Produto_Produziveis(colNodes)
    
    Exit Function
        
Erro_Carga_Arvore_Produto_Produziveis:

    Carga_Arvore_Produto_Produziveis = Err
    
    Select Case Err

        Case 65100
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142239)
        
    End Select
    
    Exit Function

End Function

Function RelOpcoes_Le_Parametros(sCodRel As String, sNome As String, colParametros As Collection) As Long
'Lê os Parametros do Relatorio para a Opção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RelOpcoes_Le_Parametros
    
    lErro = Rotina_ObterObjetoFuncao("ADM_RELOPCOES_LE_PARAMETROS", X)
    If lErro <> SUCESSO Then Error 65477
    
    RelOpcoes_Le_Parametros = X.RelOpcoes_Le_Parametros(sCodRel, sNome, colParametros)
    
    Exit Function
        
Erro_RelOpcoes_Le_Parametros:

    RelOpcoes_Le_Parametros = Err
    
    Select Case Err

        Case 65477
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142240)
        
    End Select
    
    Exit Function

End Function

Function RelOpcoes_Grava_Parametros(sCodRel As String, sNome As String, colParametros As Collection) As Long
'Lê os parametros do Relatório para Opcão
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RelOpcoes_Grava_Parametros
    
    lErro = Rotina_ObterObjetoFuncao("ADM_RELOPCOES_GRAVA_PARAMETROS", X)
    If lErro <> SUCESSO Then Error 65478
    
    RelOpcoes_Grava_Parametros = X.RelOpcoes_Grava_Parametros(sCodRel, sNome, colParametros)
    
    Exit Function
        
Erro_RelOpcoes_Grava_Parametros:

    RelOpcoes_Grava_Parametros = Err
    
    Select Case Err

        Case 65478
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142241)
        
    End Select
    
    Exit Function

End Function

Public Function RelTela_LeCampos(sCodRel As String, colCampos As ColRelTelaCampos) As Long
'carrega colecao com definicao dos campos que compoem um registro para relatorio tipo "printscreen"

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RelTela_LeCampos
    
    lErro = Rotina_ObterObjetoFuncao("ADM_RELTELA_LECAMPOS", X)
    If lErro <> SUCESSO Then Error 65479
    
    RelTela_LeCampos = X.RelTela_LeCampos(sCodRel, colCampos)
    
    Exit Function
        
Erro_RelTela_LeCampos:

    RelTela_LeCampos = Err
    
    Select Case Err

        Case 65479
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142242)
        
    End Select
    
    Exit Function

End Function

Function MenuItem_Le_Titulo(sTitulo As String, objMenuItem As ClassMenuItens) As Long
'Lê item de menu a partir do nome da tela

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MenuItem_Le_Titulo
    
    lErro = Rotina_ObterObjetoFuncao("ADM_MENUITEM_LE_TITULO", X)
    If lErro <> SUCESSO Then Error 65480
    
    MenuItem_Le_Titulo = X.MenuItem_Le_Titulo(sTitulo, objMenuItem)
    
    Exit Function
        
Erro_MenuItem_Le_Titulo:

    MenuItem_Le_Titulo = Err
    
    Select Case Err

        Case 65480
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142243)
        
    End Select
    
    Exit Function

End Function

Function VerificaRegistro_ObjetosBD(objObjetosBD As ClassObjetoBD, avParametros() As Variant) As Long

Dim X As Object, lErro As Long

On Error GoTo Erro_VerificaRegistro_ObjetosBD
    
    lErro = Rotina_ObterObjetoFuncao("CPR_VERIFICAREGISTRO_OBJETOSBD", X)
    If lErro <> SUCESSO Then gError 71420

    VerificaRegistro_ObjetosBD = X.VerificaRegistro_ObjetosBD(objObjetosBD, avParametros())

    Exit Function
    
Erro_VerificaRegistro_ObjetosBD:
    
    VerificaRegistro_ObjetosBD = gErr
    
    Select Case gErr

        Case 71420
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142244)
        
    End Select
    
    Exit Function

End Function

Function ObjetosBD_Le(objObjetoBD As ClassObjetoBD) As Long

Dim X As Object, lErro As Long

On Error GoTo Erro_ObjetosBD_Le
    
    lErro = Rotina_ObterObjetoFuncao("CPR_OBJETOSBD_LE", X)
    If lErro <> SUCESSO Then gError 71421

    ObjetosBD_Le = X.ObjetosBD_Le(objObjetoBD)

    Exit Function
    
Erro_ObjetosBD_Le:
    
    ObjetosBD_Le = gErr
    
    Select Case gErr

        Case 71421
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142245)
        
    End Select
    
    Exit Function

End Function

Function ObjetosBD_Atualiza(objObjetosBD As ClassObjetoBD) As Long

Dim X As Object, lErro As Long

On Error GoTo Erro_ObjetosBD_Atualiza
    
    lErro = Rotina_ObterObjetoFuncao("CPR_OBJETOSBD_ATUALIZA", X)
    If lErro <> SUCESSO Then gError 71422

    ObjetosBD_Atualiza = X.ObjetosBD_Atualiza(objObjetosBD)

    Exit Function
    
Erro_ObjetosBD_Atualiza:
    
    ObjetosBD_Atualiza = gErr
    
    Select Case gErr

        Case 71422
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142246)
        
    End Select
    
    Exit Function

End Function

Function Excel_Cria_Grafico(objPlanilha As ClassPlanilhaExcel) As Long
'Recebe em objPlanilha os dados que serão utilizados para montagem do gráfico
'objPlanilha deve ter uma coleção de colunas, e dentro da coleção de colunas deve existir uma coleção de células(campos)
'A planilha será montada no excel da seguinte forma: para cada coluna existente na coleção, o sistema irá preencher as células dessa coluna
'Ao preencher objPlanilha, o programador deverá informar quais as colunas que devem fazer parte do gráfico no eixo Y e qual a coluna fará parte do eixo X
'objPlanilha deve passar também algumas configurações da planilha e do gráfico como: o tipo de gráfico para cada coluna que irá compor o eixo Y,
'a forma como as séries serão montadas para plotagem (linha x coluna), título gráfico, nome do gráfico e da planilha, etc.
'Para ter um exemplo da chamada dessa função, veja a tela FluxoSintProjOcx

Dim X As Object, lErro As Long

On Error GoTo Erro_Excel_Cria_Grafico
    
    lErro = Rotina_ObterObjetoFuncao("ADM_EXCEL_CRIA_GRAFICO", X)
    If lErro <> SUCESSO Then gError 88021

    Excel_Cria_Grafico = X.Excel_Cria_Grafico(objPlanilha)

    Exit Function
    
Erro_Excel_Cria_Grafico:
    
    Excel_Cria_Grafico = gErr
    
    Select Case gErr

        Case 88021
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142247)
        
    End Select
    
    Exit Function

End Function

Function BrowseCampo_Le(sNomeTela As String, colBrowseCampo As Collection) As Long
'le todos os campos que podem ser utilizado na tela de browse e coloca os resultados na coleção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BrowseCampo_Le
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSECAMPO_LE", X)
    If lErro <> SUCESSO Then gError 89969
    
    BrowseCampo_Le = X.BrowseCampo_Le(sNomeTela, colBrowseCampo)
    
    Exit Function
    
Erro_BrowseCampo_Le:

    BrowseCampo_Le = gErr
    
    Select Case gErr

        Case 89969
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142248)
        
    End Select
    
    Exit Function

End Function

Function BrowseParamSelecao_Le(sNomeTela As String, colBrowseParamSelecao As Collection) As Long
'le os valores dos parametros de selecao relacionados a sNomeTela coloca os resultados na coleção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BrowseParamSelecao_Le
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSEPARAMSELECAO_LE", X)
    If lErro <> SUCESSO Then gError 89982
    
    BrowseParamSelecao_Le = X.BrowseParamSelecao_Le(sNomeTela, colBrowseParamSelecao)
    
    Exit Function
    
Erro_BrowseParamSelecao_Le:

    BrowseParamSelecao_Le = gErr
    
    Select Case gErr

        Case 89982
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142249)
        
    End Select
    
    Exit Function

End Function

Function BrowseUsuario_Le(objBrowseUsuario As AdmBrowseUsuario) As Long
'le os campos da tela browseusuario e coloca os resultados em objBrowseUsuario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BrowseUsuario_Le
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSEUSUARIO_LE", X)
    If lErro <> SUCESSO Then gError 89987
    
    BrowseUsuario_Le = X.BrowseUsuario_Le(objBrowseUsuario)
    
    Exit Function
    
Erro_BrowseUsuario_Le:

    BrowseUsuario_Le = gErr
    
    Select Case gErr

        Case 89987
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142250)
        
    End Select
    
    Exit Function

End Function

Function BrowseUsuario_Atualiza(objBrowseUsuario As AdmBrowseUsuario) As Long
'grava na tabela BrowseUsuario os campos contidos em objBrowseUsuario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BrowseUsuario_Atualiza
    
    lErro = Rotina_ObterObjetoFuncao("ADM_BROWSEUSUARIO_ATUALIZA", X)
    If lErro <> SUCESSO Then gError 89988
    
    BrowseUsuario_Atualiza = X.BrowseUsuario_Atualiza(objBrowseUsuario)
    
    Exit Function
    
Erro_BrowseUsuario_Atualiza:

    BrowseUsuario_Atualiza = gErr
    
    Select Case gErr

        Case 89988
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142251)
        
    End Select
    
    Exit Function

End Function

Function CF(sNomeFuncao As String, ParamArray vParams()) As Long

Dim X As Object, lErro As Long

On Error GoTo Erro_CF

    lErro = Rotina_ObterObjetoFuncao(sNomeFuncao, X)
    If lErro <> SUCESSO Then gError 89991


    If UBound(vParams) < 0 Then
        CF = CallByName(X, sNomeFuncao, VbMethod)
    ElseIf UBound(vParams) < 1 Then
        CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0))
    ElseIf UBound(vParams) < 2 Then
        CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1))
    ElseIf UBound(vParams) < 3 Then
        CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2))
    ElseIf UBound(vParams) < 4 Then
        CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3))
    ElseIf UBound(vParams) < 5 Then
        CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4))
    ElseIf UBound(vParams) < 6 Then
        CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5))
    ElseIf UBound(vParams) < 7 Then
        CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6))
    ElseIf UBound(vParams) < 8 Then
        CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7))
    ElseIf UBound(vParams) < 9 Then
        CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8))
    ElseIf UBound(vParams) < 10 Then
        CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vParams(1), vParams(2), vParams(3), vParams(4), vParams(5), vParams(6), vParams(7), vParams(8), vParams(9))
    Else
        gError 89992
    End If

    Exit Function

Erro_CF:

    CF = gErr

    Select Case gErr

        Case 89991
            lErro = Rotina_Erro(vbOKOnly, "Erro na via CF a " & sNomeFuncao, gErr)

        Case 89992
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_SUPERIOR_DE_PARAMETROS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142252)

    End Select

    Exit Function

End Function

'Function CF1(sNomeFuncao As String, Optional vp1 As Variant, Optional vP2 As Variant, Optional vP3 As Variant, Optional vP4 As Variant, Optional vP5 As Variant, Optional vP6 As Variant, Optional vP7 As Variant, Optional vP8 As Variant, Optional vP9 As Variant, Optional vP10 As Variant) As Long
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_CF1
'
'    lErro = Rotina_ObterObjetoFuncao(sNomeFuncao, X)
'    If lErro <> SUCESSO Then gError 89991
'
'    If IsMissing(vp1) < 0 Then
'        CF1 = CallByName(X, sNomeFuncao, VbMethod)
'    ElseIf IsMissing(vP2) < 0 Then
'        CF1 = CallByName(X, sNomeFuncao, VbMethod, vp1)
'    ElseIf IsMissing(vP3) < 0 Then
'        CF1 = CallByName(X, sNomeFuncao, VbMethod, vp1, vP2)
'    ElseIf IsMissing(vP4) < 0 Then
'        CF1 = CallByName(X, sNomeFuncao, VbMethod, vp1, vP2, vP3)
'    ElseIf IsMissing(vP5) < 0 Then
'        CF1 = CallByName(X, sNomeFuncao, VbMethod, vp1, vP2, vP3, vP4)
'    ElseIf IsMissing(vP6) < 0 Then
'        CF1 = CallByName(X, sNomeFuncao, VbMethod, vp1, vP2, vP3, vP4, vP5)
'    ElseIf IsMissing(vP7) < 0 Then
'        CF1 = CallByName(X, sNomeFuncao, VbMethod, vp1, vP2, vP3, vP4, vP5, vP6)
'    ElseIf IsMissing(vP8) < 0 Then
'        CF1 = CallByName(X, sNomeFuncao, VbMethod, vp1, vP2, vP3, vP4, vP5, vP6, vP7)
'    ElseIf IsMissing(vP9) < 0 Then
'        CF1 = CallByName(X, sNomeFuncao, VbMethod, vp1, vP2, vP3, vP4, vP5, vP6, vP7, vP8)
'    ElseIf IsMissing(vP10) < 0 Then
'        CF1 = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vp1, vP2, vP3, vP4, vP5, vP6, vP7, vP8, vP9)
'    Else
'        CF1 = CallByName(X, sNomeFuncao, VbMethod, vParams(0), vp1, vP2, vP3, vP4, vP5, vP6, vP7, vP8, vP9, v10)
'    End If
'
'    Exit Function
'
'Erro_CF1:
'
'    CF1 = gErr
'
'    Select Case gErr
'
'        Case 89991
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142253)
'
'    End Select
'
'    Exit Function
'
'End Function

'Function CF(sNomeFuncao As String, ParamArray vParams()) As Long
'
'Dim X As Object, lErro As Long
'Dim objAux As Object
'Dim iIndice As Integer
'
'
'On Error GoTo Erro_CF
'
'    lErro = Rotina_ObterObjetoFuncao(sNomeFuncao, X)
'    If lErro <> SUCESSO Then gError 89991
'
'    If UBound(vParams) < 0 Then
'        CF = CallByName(X, sNomeFuncao, VbMethod)
'    ElseIf UBound(vParams) < 1 Then
'        CF = CallByName(X, sNomeFuncao, VbMethod, vParams(0)) = vbInteger, aivP(0), IIf(VarType(vParams(0)) = vbLong, alvP(0), IIf(VarType(vParams(0)) = vbDouble, advP(0), IIf(VarType(vParams(0)) = vbDate, adtvP(0), IIf(VarType(vParams(0)) = vbString, asvP(0), vParams(0)))))))
'    ElseIf UBound(vParams) < 2 Then
'        CF = CallByName(X, sNomeFuncao, VbMethod, IIf(VarType(vParams(0)) = vbInteger, aivP(0), IIf(VarType(vParams(0)) = vbLong, alvP(0), IIf(VarType(vParams(0)) = vbDouble, advP(0), IIf(VarType(vParams(0)) = vbDate, adtvP(0), IIf(VarType(vParams(0)) = vbString, asvP(0), vParams(0)))))), _
'                                                  IIf(VarType(vParams(1)) = vbInteger, aivP(1), IIf(VarType(vParams(1)) = vbLong, alvP(1), IIf(VarType(vParams(1)) = vbDouble, advP(1), IIf(VarType(vParams(1)) = vbDate, adtvP(1), IIf(VarType(vParams(1)) = vbString, asvP(1), vParams(1)))))))
'    ElseIf UBound(vParams) < 3 Then
'        CF = CallByName(X, sNomeFuncao, VbMethod, IIf(VarType(vParams(0)) = vbInteger, aivP(0), IIf(VarType(vParams(0)) = vbLong, alvP(0), IIf(VarType(vParams(0)) = vbDouble, advP(0), IIf(VarType(vParams(0)) = vbDate, adtvP(0), IIf(VarType(vParams(0)) = vbString, asvP(0), vParams(0)))))), _
'                                                  IIf(VarType(vParams(1)) = vbInteger, aivP(1), IIf(VarType(vParams(1)) = vbLong, alvP(1), IIf(VarType(vParams(1)) = vbDouble, advP(1), IIf(VarType(vParams(1)) = vbDate, adtvP(1), IIf(VarType(vParams(1)) = vbString, asvP(1), vParams(1)))))), _
'                                                  IIf(VarType(vParams(2)) = vbInteger, aivP(2), IIf(VarType(vParams(2)) = vbLong, alvP(2), IIf(VarType(vParams(2)) = vbDouble, advP(2), IIf(VarType(vParams(2)) = vbDate, adtvP(2), IIf(VarType(vParams(2)) = vbString, asvP(2), vParams(2)))))))
'    ElseIf UBound(vParams) < 4 Then
'        CF = CallByName(X, sNomeFuncao, VbMethod, IIf(VarType(vParams(0)) = vbInteger, aivP(0), IIf(VarType(vParams(0)) = vbLong, alvP(0), IIf(VarType(vParams(0)) = vbDouble, advP(0), IIf(VarType(vParams(0)) = vbDate, adtvP(0), IIf(VarType(vParams(0)) = vbString, asvP(0), vParams(0)))))), _
'                                                  IIf(VarType(vParams(1)) = vbInteger, aivP(1), IIf(VarType(vParams(1)) = vbLong, alvP(1), IIf(VarType(vParams(1)) = vbDouble, advP(1), IIf(VarType(vParams(1)) = vbDate, adtvP(1), IIf(VarType(vParams(1)) = vbString, asvP(1), vParams(1)))))), _
'                                                  IIf(VarType(vParams(2)) = vbInteger, aivP(2), IIf(VarType(vParams(2)) = vbLong, alvP(2), IIf(VarType(vParams(2)) = vbDouble, advP(2), IIf(VarType(vParams(2)) = vbDate, adtvP(2), IIf(VarType(vParams(2)) = vbString, asvP(2), vParams(2)))))), _
'                                                  IIf(VarType(vParams(3)) = vbInteger, aivP(3), IIf(VarType(vParams(3)) = vbLong, alvP(3), IIf(VarType(vParams(3)) = vbDouble, advP(3), IIf(VarType(vParams(3)) = vbDate, adtvP(3), IIf(VarType(vParams(3)) = vbString, asvP(3), vParams(3)))))))
'    ElseIf UBound(vParams) < 5 Then
'        CF = CallByName(X, sNomeFuncao, VbMethod, IIf(VarType(vParams(0)) = vbInteger, aivP(0), IIf(VarType(vParams(0)) = vbLong, alvP(0), IIf(VarType(vParams(0)) = vbDouble, advP(0), IIf(VarType(vParams(0)) = vbDate, adtvP(0), IIf(VarType(vParams(0)) = vbString, asvP(0), vParams(0)))))), _
'                                                  IIf(VarType(vParams(1)) = vbInteger, aivP(1), IIf(VarType(vParams(1)) = vbLong, alvP(1), IIf(VarType(vParams(1)) = vbDouble, advP(1), IIf(VarType(vParams(1)) = vbDate, adtvP(1), IIf(VarType(vParams(1)) = vbString, asvP(1), vParams(1)))))), _
'                                                  IIf(VarType(vParams(2)) = vbInteger, aivP(2), IIf(VarType(vParams(2)) = vbLong, alvP(2), IIf(VarType(vParams(2)) = vbDouble, advP(2), IIf(VarType(vParams(2)) = vbDate, adtvP(2), IIf(VarType(vParams(2)) = vbString, asvP(2), vParams(2)))))), _
'                                                  IIf(VarType(vParams(3)) = vbInteger, aivP(3), IIf(VarType(vParams(3)) = vbLong, alvP(3), IIf(VarType(vParams(3)) = vbDouble, advP(3), IIf(VarType(vParams(3)) = vbDate, adtvP(3), IIf(VarType(vParams(3)) = vbString, asvP(3), vParams(3)))))), _
'                                                  IIf(VarType(vParams(4)) = vbInteger, aivP(4), IIf(VarType(vParams(4)) = vbLong, alvP(4), IIf(VarType(vParams(4)) = vbDouble, advP(4), IIf(VarType(vParams(4)) = vbDate, adtvP(4), IIf(VarType(vParams(4)) = vbString, asvP(4), vParams(4)))))))
'    ElseIf UBound(vParams) < 6 Then
'        CF = CallByName(X, sNomeFuncao, VbMethod, IIf(VarType(vParams(0)) = vbInteger, aivP(0), IIf(VarType(vParams(0)) = vbLong, alvP(0), IIf(VarType(vParams(0)) = vbDouble, advP(0), IIf(VarType(vParams(0)) = vbDate, adtvP(0), IIf(VarType(vParams(0)) = vbString, asvP(0), vParams(0)))))), _
'                                                  IIf(VarType(vParams(1)) = vbInteger, aivP(1), IIf(VarType(vParams(1)) = vbLong, alvP(1), IIf(VarType(vParams(1)) = vbDouble, advP(1), IIf(VarType(vParams(1)) = vbDate, adtvP(1), IIf(VarType(vParams(1)) = vbString, asvP(1), vParams(1)))))), _
'                                                  IIf(VarType(vParams(2)) = vbInteger, aivP(2), IIf(VarType(vParams(2)) = vbLong, alvP(2), IIf(VarType(vParams(2)) = vbDouble, advP(2), IIf(VarType(vParams(2)) = vbDate, adtvP(2), IIf(VarType(vParams(2)) = vbString, asvP(2), vParams(2)))))), _
'                                                  IIf(VarType(vParams(3)) = vbInteger, aivP(3), IIf(VarType(vParams(3)) = vbLong, alvP(3), IIf(VarType(vParams(3)) = vbDouble, advP(3), IIf(VarType(vParams(3)) = vbDate, adtvP(3), IIf(VarType(vParams(3)) = vbString, asvP(3), vParams(3)))))), _
'                                                  IIf(VarType(vParams(4)) = vbInteger, aivP(4), IIf(VarType(vParams(4)) = vbLong, alvP(4), IIf(VarType(vParams(4)) = vbDouble, advP(4), IIf(VarType(vParams(4)) = vbDate, adtvP(4), IIf(VarType(vParams(4)) = vbString, asvP(4), vParams(4)))))), _
'                                                  IIf(VarType(vParams(5)) = vbInteger, aivP(5), IIf(VarType(vParams(5)) = vbLong, alvP(5), IIf(VarType(vParams(5)) = vbDouble, advP(5), IIf(VarType(vParams(5)) = vbDate, adtvP(5), IIf(VarType(vParams(5)) = vbString, asvP(5), vParams(5)))))))
'    ElseIf UBound(vParams) < 7 Then
'        CF = CallByName(X, sNomeFuncao, VbMethod, IIf(VarType(vParams(0)) = vbInteger, aivP(0), IIf(VarType(vParams(0)) = vbLong, alvP(0), IIf(VarType(vParams(0)) = vbDouble, advP(0), IIf(VarType(vParams(0)) = vbDate, adtvP(0), IIf(VarType(vParams(0)) = vbString, asvP(0), vParams(0)))))), _
'                                                  IIf(VarType(vParams(1)) = vbInteger, aivP(1), IIf(VarType(vParams(1)) = vbLong, alvP(1), IIf(VarType(vParams(1)) = vbDouble, advP(1), IIf(VarType(vParams(1)) = vbDate, adtvP(1), IIf(VarType(vParams(1)) = vbString, asvP(1), vParams(1)))))), _
'                                                  IIf(VarType(vParams(2)) = vbInteger, aivP(2), IIf(VarType(vParams(2)) = vbLong, alvP(2), IIf(VarType(vParams(2)) = vbDouble, advP(2), IIf(VarType(vParams(2)) = vbDate, adtvP(2), IIf(VarType(vParams(2)) = vbString, asvP(2), vParams(2)))))), _
'                                                  IIf(VarType(vParams(3)) = vbInteger, aivP(3), IIf(VarType(vParams(3)) = vbLong, alvP(3), IIf(VarType(vParams(3)) = vbDouble, advP(3), IIf(VarType(vParams(3)) = vbDate, adtvP(3), IIf(VarType(vParams(3)) = vbString, asvP(3), vParams(3)))))), _
'                                                  IIf(VarType(vParams(4)) = vbInteger, aivP(4), IIf(VarType(vParams(4)) = vbLong, alvP(4), IIf(VarType(vParams(4)) = vbDouble, advP(4), IIf(VarType(vParams(4)) = vbDate, adtvP(4), IIf(VarType(vParams(4)) = vbString, asvP(4), vParams(4)))))), _
'                                                  IIf(VarType(vParams(5)) = vbInteger, aivP(5), IIf(VarType(vParams(5)) = vbLong, alvP(5), IIf(VarType(vParams(5)) = vbDouble, advP(5), IIf(VarType(vParams(5)) = vbDate, adtvP(5), IIf(VarType(vParams(5)) = vbString, asvP(5), vParams(5)))))), _
'                                                  IIf(VarType(vParams(6)) = vbInteger, aivP(6), IIf(VarType(vParams(6)) = vbLong, alvP(6), IIf(VarType(vParams(6)) = vbDouble, advP(6), IIf(VarType(vParams(6)) = vbDate, adtvP(6), IIf(VarType(vParams(6)) = vbString, asvP(6), vParams(6)))))))
'    ElseIf UBound(vParams) < 8 Then
'        CF = CallByName(X, sNomeFuncao, VbMethod, IIf(VarType(vParams(0)) = vbInteger, aivP(0), IIf(VarType(vParams(0)) = vbLong, alvP(0), IIf(VarType(vParams(0)) = vbDouble, advP(0), IIf(VarType(vParams(0)) = vbDate, adtvP(0), IIf(VarType(vParams(0)) = vbString, asvP(0), vParams(0)))))), _
'                                                  IIf(VarType(vParams(1)) = vbInteger, aivP(1), IIf(VarType(vParams(1)) = vbLong, alvP(1), IIf(VarType(vParams(1)) = vbDouble, advP(1), IIf(VarType(vParams(1)) = vbDate, adtvP(1), IIf(VarType(vParams(1)) = vbString, asvP(1), vParams(1)))))), _
'                                                  IIf(VarType(vParams(2)) = vbInteger, aivP(2), IIf(VarType(vParams(2)) = vbLong, alvP(2), IIf(VarType(vParams(2)) = vbDouble, advP(2), IIf(VarType(vParams(2)) = vbDate, adtvP(2), IIf(VarType(vParams(2)) = vbString, asvP(2), vParams(2)))))), _
'                                                  IIf(VarType(vParams(3)) = vbInteger, aivP(3), IIf(VarType(vParams(3)) = vbLong, alvP(3), IIf(VarType(vParams(3)) = vbDouble, advP(3), IIf(VarType(vParams(3)) = vbDate, adtvP(3), IIf(VarType(vParams(3)) = vbString, asvP(3), vParams(3)))))), _
'                                                  IIf(VarType(vParams(4)) = vbInteger, aivP(4), IIf(VarType(vParams(4)) = vbLong, alvP(4), IIf(VarType(vParams(4)) = vbDouble, advP(4), IIf(VarType(vParams(4)) = vbDate, adtvP(4), IIf(VarType(vParams(4)) = vbString, asvP(4), vParams(4)))))), _
'                                                  IIf(VarType(vParams(5)) = vbInteger, aivP(5), IIf(VarType(vParams(5)) = vbLong, alvP(5), IIf(VarType(vParams(5)) = vbDouble, advP(5), IIf(VarType(vParams(5)) = vbDate, adtvP(5), IIf(VarType(vParams(5)) = vbString, asvP(5), vParams(5)))))), _
'                                                  IIf(VarType(vParams(6)) = vbInteger, aivP(6), IIf(VarType(vParams(6)) = vbLong, alvP(6), IIf(VarType(vParams(6)) = vbDouble, advP(6), IIf(VarType(vParams(6)) = vbDate, adtvP(6), IIf(VarType(vParams(6)) = vbString, asvP(6), vParams(6)))))), _
'                                                  IIf(VarType(vParams(7)) = vbInteger, aivP(7), IIf(VarType(vParams(7)) = vbLong, alvP(7), IIf(VarType(vParams(7)) = vbDouble, advP(7), IIf(VarType(vParams(7)) = vbDate, adtvP(7), IIf(VarType(vParams(7)) = vbString, asvP(7), vParams(7)))))))
'    ElseIf UBound(vParams) < 9 Then
'        CF = CallByName(X, sNomeFuncao, VbMethod, IIf(VarType(vParams(0)) = vbInteger, aivP(0), IIf(VarType(vParams(0)) = vbLong, alvP(0), IIf(VarType(vParams(0)) = vbDouble, advP(0), IIf(VarType(vParams(0)) = vbDate, adtvP(0), IIf(VarType(vParams(0)) = vbString, asvP(0), vParams(0)))))), _
'                                                  IIf(VarType(vParams(1)) = vbInteger, aivP(1), IIf(VarType(vParams(1)) = vbLong, alvP(1), IIf(VarType(vParams(1)) = vbDouble, advP(1), IIf(VarType(vParams(1)) = vbDate, adtvP(1), IIf(VarType(vParams(1)) = vbString, asvP(1), vParams(1)))))), _
'                                                  IIf(VarType(vParams(2)) = vbInteger, aivP(2), IIf(VarType(vParams(2)) = vbLong, alvP(2), IIf(VarType(vParams(2)) = vbDouble, advP(2), IIf(VarType(vParams(2)) = vbDate, adtvP(2), IIf(VarType(vParams(2)) = vbString, asvP(2), vParams(2)))))), _
'                                                  IIf(VarType(vParams(3)) = vbInteger, aivP(3), IIf(VarType(vParams(3)) = vbLong, alvP(3), IIf(VarType(vParams(3)) = vbDouble, advP(3), IIf(VarType(vParams(3)) = vbDate, adtvP(3), IIf(VarType(vParams(3)) = vbString, asvP(3), vParams(3)))))), _
'                                                  IIf(VarType(vParams(4)) = vbInteger, aivP(4), IIf(VarType(vParams(4)) = vbLong, alvP(4), IIf(VarType(vParams(4)) = vbDouble, advP(4), IIf(VarType(vParams(4)) = vbDate, adtvP(4), IIf(VarType(vParams(4)) = vbString, asvP(4), vParams(4)))))), _
'                                                  IIf(VarType(vParams(5)) = vbInteger, aivP(5), IIf(VarType(vParams(5)) = vbLong, alvP(5), IIf(VarType(vParams(5)) = vbDouble, advP(5), IIf(VarType(vParams(5)) = vbDate, adtvP(5), IIf(VarType(vParams(5)) = vbString, asvP(5), vParams(5)))))), _
'                                                  IIf(VarType(vParams(6)) = vbInteger, aivP(6), IIf(VarType(vParams(6)) = vbLong, alvP(6), IIf(VarType(vParams(6)) = vbDouble, advP(6), IIf(VarType(vParams(6)) = vbDate, adtvP(6), IIf(VarType(vParams(6)) = vbString, asvP(6), vParams(6)))))), _
'                                                  IIf(VarType(vParams(7)) = vbInteger, aivP(7), IIf(VarType(vParams(7)) = vbLong, alvP(7), IIf(VarType(vParams(7)) = vbDouble, advP(7), IIf(VarType(vParams(7)) = vbDate, adtvP(7), IIf(VarType(vParams(7)) = vbString, asvP(7), vParams(7)))))), _
'                                                  IIf(VarType(vParams(8)) = vbInteger, aivP(8), IIf(VarType(vParams(8)) = vbLong, alvP(8), IIf(VarType(vParams(8)) = vbDouble, advP(8), IIf(VarType(vParams(8)) = vbDate, adtvP(8), IIf(VarType(vParams(8)) = vbString, asvP(8), vParams(8)))))))
'    ElseIf UBound(vParams) < 10 Then
'        CF = CallByName(X, sNomeFuncao, VbMethod, IIf(VarType(vParams(0)) = vbInteger, aivP(0), IIf(VarType(vParams(0)) = vbLong, alvP(0), IIf(VarType(vParams(0)) = vbDouble, advP(0), IIf(VarType(vParams(0)) = vbDate, adtvP(0), IIf(VarType(vParams(0)) = vbString, asvP(0), vParams(0)))))), _
'                                                  IIf(VarType(vParams(1)) = vbInteger, aivP(1), IIf(VarType(vParams(1)) = vbLong, alvP(1), IIf(VarType(vParams(1)) = vbDouble, advP(1), IIf(VarType(vParams(1)) = vbDate, adtvP(1), IIf(VarType(vParams(1)) = vbString, asvP(1), vParams(1)))))), _
'                                                  IIf(VarType(vParams(2)) = vbInteger, aivP(2), IIf(VarType(vParams(2)) = vbLong, alvP(2), IIf(VarType(vParams(2)) = vbDouble, advP(2), IIf(VarType(vParams(2)) = vbDate, adtvP(2), IIf(VarType(vParams(2)) = vbString, asvP(2), vParams(2)))))), _
'                                                  IIf(VarType(vParams(3)) = vbInteger, aivP(3), IIf(VarType(vParams(3)) = vbLong, alvP(3), IIf(VarType(vParams(3)) = vbDouble, advP(3), IIf(VarType(vParams(3)) = vbDate, adtvP(3), IIf(VarType(vParams(3)) = vbString, asvP(3), vParams(3)))))), _
'                                                  IIf(VarType(vParams(4)) = vbInteger, aivP(4), IIf(VarType(vParams(4)) = vbLong, alvP(4), IIf(VarType(vParams(4)) = vbDouble, advP(4), IIf(VarType(vParams(4)) = vbDate, adtvP(4), IIf(VarType(vParams(4)) = vbString, asvP(4), vParams(4)))))), _
'                                                  IIf(VarType(vParams(5)) = vbInteger, aivP(5), IIf(VarType(vParams(5)) = vbLong, alvP(5), IIf(VarType(vParams(5)) = vbDouble, advP(5), IIf(VarType(vParams(5)) = vbDate, adtvP(5), IIf(VarType(vParams(5)) = vbString, asvP(5), vParams(5)))))), _
'                                                  IIf(VarType(vParams(6)) = vbInteger, aivP(6), IIf(VarType(vParams(6)) = vbLong, alvP(6), IIf(VarType(vParams(6)) = vbDouble, advP(6), IIf(VarType(vParams(6)) = vbDate, adtvP(6), IIf(VarType(vParams(6)) = vbString, asvP(6), vParams(6)))))), _
'                                                  IIf(VarType(vParams(7)) = vbInteger, aivP(7), IIf(VarType(vParams(7)) = vbLong, alvP(7), IIf(VarType(vParams(7)) = vbDouble, advP(7), IIf(VarType(vParams(7)) = vbDate, adtvP(7), IIf(VarType(vParams(7)) = vbString, asvP(7), vParams(7)))))), _
'                                                  IIf(VarType(vParams(8)) = vbInteger, aivP(8), IIf(VarType(vParams(8)) = vbLong, alvP(8), IIf(VarType(vParams(8)) = vbDouble, advP(8), IIf(VarType(vParams(8)) = vbDate, adtvP(8), IIf(VarType(vParams(8)) = vbString, asvP(8), vParams(8)))))), _
'                                                  IIf(VarType(vParams(9)) = vbInteger, aivP(9), IIf(VarType(vParams(9)) = vbLong, alvP(9), IIf(VarType(vParams(9)) = vbDouble, advP(9), IIf(VarType(vParams(9)) = vbDate, adtvP(9), IIf(VarType(vParams(9)) = vbString, asvP(9), vParams(9)))))))
'    Else
'        gError 89991
'
'    End If
'
'    Exit Function
'
'Erro_CF:
'
'    CF = gErr
'
'    Select Case gErr
'
'        Case 89991
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142254)
'
'    End Select
'
'    Exit Function
'
'End Function


