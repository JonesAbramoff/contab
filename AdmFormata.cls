VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmFormata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const EXCEL_XLCATEGORY = 1

'Function MesNome(iMes As Integer, sMes As String) As Long
''Entra com o número do mês. Retorna-o por extenso.
'
'Dim lErro As Long
'
'On Error GoTo Erro_MesNome
'
'    'Seleciona o mês que corresponde a iMes
'    Select Case iMes
'
'        Case 1
'            sMes = "Janeiro"
'
'        Case 2
'            sMes = "Fevereiro"
'
'        Case 3
'            sMes = "Março"
'
'        Case 4
'            sMes = "Abril"
'
'        Case 5
'            sMes = "Maio"
'
'        Case 6
'            sMes = "Junho"
'
'        Case 7
'            sMes = "Julho"
'
'        Case 8
'            sMes = "Agosto"
'
'        Case 9
'            sMes = "Setembro"
'
'        Case 10
'            sMes = "Outubro"
'
'        Case 11
'            sMes = "Novembro"
'
'        Case 12
'            sMes = "Dezembro"
'
'        Case Else
'            Error 25208
'
'    End Select
'
'    MesNome = SUCESSO
'
'    Exit Function
'
'Erro_MesNome:
'
'    MesNome = Err
'
'    Select Case Err
'
'        Case 25208
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_NAO_MES", Err, iMes)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142316)
'
'    End Select
'
'    Exit Function
'
'End Function
'Function MesNumero(sMes As String, iMes As Integer) As Long
''Entra com o mês por extenso. Retorna-o como um inteiro.
'
'Dim lErro As Long
'
'On Error GoTo Erro_MesNumero
'
'    'Seleciona o mês que corresponde a sMes
'    Select Case sMes
'
'        Case "Janeiro"
'            iMes = 1
'
'        Case "Fevereiro"
'            iMes = 2
'
'        Case "Março"
'            iMes = 3
'
'        Case "Abril"
'            iMes = 4
'
'        Case "Maio"
'            iMes = 5
'
'        Case "Junho"
'            iMes = 6
'
'        Case "Julho"
'            iMes = 7
'
'        Case "Agosto"
'            iMes = 8
'
'        Case "Setembro"
'            iMes = 9
'
'        Case "Outubro"
'            iMes = 10
'
'        Case "Novembro"
'            iMes = 11
'
'        Case "Dezembro"
'            iMes = 12
'
'        Case Else
'            Error 25209
'
'    End Select
'
'    MesNumero = SUCESSO
'
'    Exit Function
'
'Erro_MesNumero:
'
'    MesNumero = Err
'
'    Select Case Err
'
'        Case 25209
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_STRING_NAO_MES", Err, sMes)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142317)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function StrParaDate(sTexto As String) As Date
''retorna texto como data
'
'    If IsDate(sTexto) Then
'        StrParaDate = CDate(sTexto)
'    Else
'        StrParaDate = DATA_NULA
'    End If
'
'End Function
'
'Function StrParaDbl(sTexto As String) As Double
''retorna sTexto como double
'
'    If Len(Trim(sTexto)) = 0 Then
'        StrParaDbl = 0
'    Else
'        StrParaDbl = CDbl(sTexto)
'    End If
'
'End Function
'Function PercentParaDbl(sTexto As String) As Double
''retorna percentagem sTexto como double
''sTexto deve ser número finalizado por %
'
'    If Len(Trim(sTexto)) = 0 Then
'        PercentParaDbl = 0
'    ElseIf Not IsNumeric(Left(sTexto, Len(sTexto) - 1)) Then
'        PercentParaDbl = 0
'    ElseIf Right(sTexto, 1) <> "%" Then
'        PercentParaDbl = 0
'    Else
'        PercentParaDbl = CDbl(Format(sTexto, "General Number"))
'    End If
'
'End Function
'
'Function StrParaInt(sTexto As String) As Integer
''retorna sTexto como Integer
'
'    If Len(Trim(sTexto)) = 0 Then
'        StrParaInt = 0
'    Else
'        StrParaInt = CInt(sTexto)
'    End If
'
'End Function
'
'Function StrParaLong(sTexto As String) As Long
''retorna sTexto como Long
'
'    If Len(Trim(sTexto)) = 0 Then
'        StrParaLong = 0
'    Else
'        StrParaLong = CLng(sTexto)
'    End If
'
'End Function
'
'Sub DateParaMasked(objMasked As Object, dtData As Date)
''preenche MaskedEdit com Date
'
'    If dtData <> DATA_NULA Then
'        objMasked.PromptInclude = False
'        objMasked.Text = Format(dtData, "dd/mm/yy")
'        objMasked.PromptInclude = True
'    Else
'        objMasked.PromptInclude = False
'        objMasked.Text = ""
'        objMasked.PromptInclude = True
'    End If
'
'End Sub
'
'Function MaskedParaDate(objMasked As Object) As Date
''retorna objMasked.Text como Date
'
'    If Len(Trim(objMasked.ClipText)) = 0 Then
'        MaskedParaDate = DATA_NULA
'    Else
'        MaskedParaDate = CDate(objMasked.FormattedText)
'    End If
'
'End Function
'
'Function StrParaMasked(objMasked As Object, sString As String)
''coloca sString em MaskedEdit verificando promptinclude
'
'    If objMasked.PromptInclude = True Then
'
'        objMasked.PromptInclude = False
'        objMasked.Text = sString
'        objMasked.PromptInclude = True
'
'    Else
'
'        objMasked.Text = sString
'
'    End If
'
'End Function
'
'Sub Combo_Seleciona_ItemData(objComboBox As Object, ByVal lItemData As Long)
''seleciona o item da combo que tem ItemData igual ao passado por parametro
'
'    Call ListOuCombo_Seleciona_ItemData(objComboBox, lItemData)
'
'End Sub
'
'Sub List_Seleciona_ItemData(objListBox As Object, ByVal lItemData As Long)
''seleciona o item da listbox que tem ItemData igual ao passado por parametro
'
'    Call ListOuCombo_Seleciona_ItemData(objListBox, lItemData)
'
'End Sub
'
'Private Sub ListOuCombo_Seleciona_ItemData(objBox As Object, ByVal lItemData As Long)
''seleciona o item da combo ou listbox
'Dim lErro As Long, iIndice As Integer
'On Error GoTo Erro_ListOuCombo_Seleciona_ItemData
'
'    'procura um item com o iTemData informado
'    For iIndice = 0 To objBox.ListCount - 1
'
'        If objBox.ItemData(iIndice) = lItemData Then
'            objBox.ListIndex = iIndice
'            Exit For
'        End If
'
'    Next
'
'    Exit Sub
'
'Erro_ListOuCombo_Seleciona_ItemData:
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142318)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Function ListOuCombo_Item_Igual(objBox As Object, sString As String) As Long
''Seleciona ítem na listobox ou ComboBox referente a sString passada por parametro
''ATENÇÃO: Esta função ao selecionar ativa o ClickEvent
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim iSucesso As Integer
'
'On Error GoTo Error_ListOuCombo_Item_Igual
'
'    iSucesso = 0
'
'    'percorre todos os elementos da Box
'    For iIndice = 0 To objBox.ListCount - 1
'
'        'compara se codigo ja existe na Box
'        If objBox.List(iIndice) = sString Then
'
'            iSucesso = 1
'
'            'Seleciona o ítem
'            objBox.ListIndex = iIndice
'            Exit For
'
'        End If
'
'    Next
'
'    If iSucesso <> 1 Then Error 12253
'
'    ListOuCombo_Item_Igual = SUCESSO
'
'Exit Function
'
'Error_ListOuCombo_Item_Igual:
'
'    ListOuCombo_Item_Igual = Err
'
'    Select Case Err
'
'        Case 12253 'nao encontrou
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142319)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function ListOuCombo_Item_Igual_Remove(objBox As Object, sString As String) As Long
''Remove ítem na ListBox ou ComboBox referente a sString passada por parametro
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim iSucesso As Integer
'
'On Error GoTo Error_ListOuCombo_Item_Igual_Remove
'
'    iSucesso = 0
'
'    'percorre todos os elementos da ComboBox
'    For iIndice = 0 To objBox.ListCount - 1
'
'        'compara se codigo ja existe na ComboBox
'        If objBox.List(iIndice) = sString Then
'
'            iSucesso = 1
'
'            'remove o ítem da Combo
'            objBox.RemoveItem (iIndice)
'            Exit For
'
'        End If
'
'    Next
'
'    If iSucesso <> 1 Then Error 19370
'
'    ListOuCombo_Item_Igual_Remove = SUCESSO
'
'Exit Function
'
'Error_ListOuCombo_Item_Igual_Remove:
'
'    ListOuCombo_Item_Igual_Remove = Err
'
'    Select Case Err
'
'        Case 19370 'nao encontrou
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142320)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function SeguraFoco(objControle As Object, Frames As Object, iFrameAtual As Integer, Opcao As Object) As Long
''Segura o foco em objControle quando o Controle está em uma Frame de Frames dentro do TabStrip Opcao
''Frames deve se referir às Frames controladas pelo TabStrip
'
'Dim lErro As Long
'Dim objFrame As Frame  'Será a Frame (ligada a TabStrip) que contém o Controle
'
'On Error GoTo Erro_SeguraFoco
'
'    'Testa a profundidade do Frame em relação ao Controle
'    If objControle.Container.Container <> "Frame" Then
'        Set objFrame = objControle.Container
'    ElseIf objControle.Container.Container.Container <> "Frame" Then
'        Set objFrame = objControle.Container.Container
'    Else
'        Set objFrame = objControle.Container.Container.Container
'    End If
'
'    'Se FrameAtual não é o do controle
'    If objFrame.Index <> iFrameAtual Then
'
'       'Torna Frame do controle visível
'        objFrame.Visible = True
'        Frames(iFrameAtual).Visible = False
'        iFrameAtual = objFrame.Index
'
'        'Torna Tab da Frame do Controle o Tab selecionado
'        Set Opcao.SelectedItem = Opcao.Tabs(objFrame.Index)
'
'    End If
'
'    'Segura o foco
'    objControle.SetFocus
'
'    SeguraFoco = SUCESSO
'
'    Exit Function
'
'Erro_SeguraFoco:
'
'    SeguraFoco = Err
'
'    lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142321)
'
'    Exit Function
'
'End Function
'
'
'Function Data_Aumenta(sData As String) As Long
'
'Dim lErro As Long
'Dim dtData As Date
'
'On Error GoTo Erro_Data_Aumenta
'
'    'Se está vazia deixa como está
'    If sData <> "  /  /  " Then
'
'        'Critica data
'        If Not IsDate(sData) Then Error 6174
'
'        'Aumenta data de um dia
'        dtData = CDate(sData)
'        dtData = DateAdd("d", 1, dtData)
'        sData = Format(dtData, "dd/mm/yy")
'
'    End If
'
'
'    Data_Aumenta = SUCESSO
'
'    Exit Function
'
'Erro_Data_Aumenta:
'
'    Data_Aumenta = Err
'
'    Select Case Err
'
'        Case 6174 'Vazio porque o erro já foi tratado
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142322)
'
'    End Select
'
'    Exit Function
'
'End Function
'Function Data_Diminui(sData As String) As Long
'
'Dim lErro As Long
'Dim dtData As Date
'
'On Error GoTo Erro_Data_Diminui
'
'    'Se está vazia deixa como está
'    If sData <> "  /  /  " Then
'
'        'Critica data
'        If Not IsDate(sData) Then Error 6175
'
'        'Aumenta data de um dia
'        dtData = CDate(sData)
'        dtData = DateAdd("d", -1, dtData)
'        sData = Format(dtData, "dd/mm/yy")
'
'    End If
'
'    Data_Diminui = SUCESSO
'
'    Exit Function
'
'Erro_Data_Diminui:
'
'    Data_Diminui = Err
'
'    Select Case Err
'
'        Case 6175 'Vazio porque o erro já foi tratado
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142323)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Data_Up_Down_Click(MaskEdData As Object, ByVal iOpcao As Integer) As Long
''recebe o MaskEdBox de uma Data e a opção ( AUMENTA_DATA ou DIMINUI_DATA )
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_Data_Up_Down_Click
'
'    MaskEdData.SetFocus
'
'    'verifica se a data está vazia
'    If Len(MaskEdData.ClipText) > 0 Then
'
'        sData = MaskEdData.Text
'
'        If iOpcao = AUMENTA_DATA Then
'
'            lErro = Data_Aumenta(sData)
'            If lErro <> SUCESSO Then Error 13064
'
'        Else
'
'            lErro = Data_Diminui(sData)
'            If lErro <> SUCESSO Then Error 13041
'
'        End If
'
'        'exibe o novo valor da data
'        MaskEdData.PromptInclude = False
'        MaskEdData.Text = sData
'        MaskEdData.PromptInclude = True
'
'    End If
'
'    Data_Up_Down_Click = SUCESSO
'
'    Exit Function
'
'Erro_Data_Up_Down_Click:
'
'    Data_Up_Down_Click = Err
'
'    Select Case Err
'
'        Case 13041
'
'        Case 13064
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142324)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Codigo_Extrai(sString As String) As Integer
'
'Dim iPosicaoSeparador As Integer, sBuffer As String, dCodigo As Double
'On Error GoTo Erro_Codigo_Extrai
'
'    iPosicaoSeparador = InStr(sString, SEPARADOR)
'    If iPosicaoSeparador > 1 Then
'        sBuffer = Mid(sString, 1, iPosicaoSeparador - 1)
'    Else
'        sBuffer = sString
'    End If
'
'    If IsNumeric(sBuffer) = False Then Error 7717
'
'    Codigo_Extrai = CInt(sBuffer)
'
'    Exit Function
'
'Erro_Codigo_Extrai:
'
'    Codigo_Extrai = 0
'
'    Exit Function
'
'End Function
'
'Function LCodigo_Extrai(sString As String) As Long
'
'Dim iPosicaoSeparador As Integer, sBuffer As String, dCodigo As Double
'On Error GoTo Erro_LCodigo_Extrai
'
'    iPosicaoSeparador = InStr(sString, SEPARADOR)
'    If iPosicaoSeparador > 1 Then
'        sBuffer = Mid(sString, 1, iPosicaoSeparador - 1)
'    Else
'        sBuffer = sString
'    End If
'
'    If IsNumeric(sBuffer) = False Then Error 7718
'
'    LCodigo_Extrai = CLng(sBuffer)
'
'    Exit Function
'
'Erro_LCodigo_Extrai:
'
'    LCodigo_Extrai = 0
'
'    Exit Function
'
'End Function
'
'Function SCodigo_Extrai(sString As String) As String
'
'Dim iPosicaoSeparador As Integer, sBuffer As String
'On Error GoTo Erro_SCodigo_Extrai
'
'    iPosicaoSeparador = InStr(sString, SEPARADOR)
'    If iPosicaoSeparador > 1 Then
'        sBuffer = Mid(sString, 1, iPosicaoSeparador - 1)
'    Else
'        sBuffer = sString
'    End If
'
'    SCodigo_Extrai = sBuffer
'
'    Exit Function
'
'Erro_SCodigo_Extrai:
'
'    SCodigo_Extrai = ""
'
'    Exit Function
'
'End Function
'
'Function Combo_Item_Igual_CI(objComboBox As Object)
''Seleciona ítem na ComboBox referente ao objComboBox.text passado por parametro
''Seleciona independente de ser digitada letra maiuscula ou letra minuscula
''ATENÇÃO: Esta função ao selecionar ativa o ClickEvent
'
'    Combo_Item_Igual_CI = ListOuCombo_Item_Igual_CI(objComboBox, Trim(objComboBox.Text))
'
'End Function
'
'Private Function ListOuCombo_Item_Igual_CI(objBox As Object, sString As String) As Long
''Seleciona ítem na listobox ou ComboBox referente a sString passada por parametro
''Seleciona independente de ser digitada letra maiuscula ou letra minuscula
''ATENÇÃO: Esta função ao selecionar ativa o ClickEvent
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim iSucesso As Integer
'
'On Error GoTo Error_ListOuCombo_Item_Igual_CI
'
'    iSucesso = 0
'
'    'percorre todos os elementos da Box
'    For iIndice = 0 To objBox.ListCount - 1
'
'        'compara se codigo ja existe na Box
'        If UCase(objBox.List(iIndice)) = UCase(sString) Then
'
'            iSucesso = 1
'
'            'Seleciona o ítem
'            objBox.ListIndex = iIndice
'            Exit For
'
'        End If
'
'    Next
'
'    If iSucesso <> 1 Then Error 58583
'
'    ListOuCombo_Item_Igual_CI = SUCESSO
'
'    Exit Function
'
'Error_ListOuCombo_Item_Igual_CI:
'
'    ListOuCombo_Item_Igual_CI = Err
'
'    Select Case Err
'
'        Case 58583 'nao encontrou
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142325)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Combo_Item_Igual(objComboBox As Object) As Long
''Seleciona ítem na ComboBox referente ao objComboBox.text passado por parametro
''ATENÇÃO: Esta função ao selecionar ativa o ClickEvent
'
'    Combo_Item_Igual = ListOuCombo_Item_Igual(objComboBox, Trim(objComboBox.Text))
'
'End Function
'
'Function Combo_Item_Igual_Remove(objComboBox As Object) As Long
''Remove ítem na ComboBox referente ao objComboBox.text passado por parametro
''ATENÇÃO: Esta função ao selecionar ativa o ClickEvent
'
'    Combo_Item_Igual_Remove = ListOuCombo_Item_Igual_Remove(objComboBox, Trim(objComboBox.Text))
'
'End Function
'
'Function List_Item_Igual(objListBox As Object, sString As String) As Long
''Seleciona ítem na ListBox referente a sString passada por parametro
''ATENÇÃO: Esta função ao selecionar ativa o ClickEvent
'
'    List_Item_Igual = ListOuCombo_Item_Igual(objListBox, sString)
'
'End Function
'
'Function List_Item_Igual_Remove(objListBox As Object, sString As String) As Long
''remove ítem na ListBox referente a sString passada por parametro
'
'    List_Item_Igual_Remove = ListOuCombo_Item_Igual_Remove(objListBox, sString)
'
'End Function
'
''Seleciona o ítem da Combo correspondente ao Text que pode ser numérico ou string
'Function Combo_Seleciona(objComboBox As Object, iCodigo As Integer) As Long
'
'Dim objCombo As ComboBox
'Dim iNumerico As Integer
'Dim iEncontrado As Integer
'Dim lErro As Long
'Dim dNumero As Double
'Dim iIndice As Integer
'Dim sString As String
'Dim iPosicao As String
'
'On Error GoTo Erro_Combo_Seleciona
'
'    Set objCombo = objComboBox
'
'    'Verifica se é o código numérico
'    iNumerico = IsNumeric(objCombo.Text)
'
'    If iNumerico Then
'
'        dNumero = CDbl(objCombo.Text)
'
'    Else  'Não é número puro
'        'Determina existência e posição do caracter separador
'        iPosicao = InStr(objCombo.Text, SEPARADOR)
'
'        If iPosicao = 0 Then    'Não existe separador
'
'            iNumerico = False
'
'        Else                    'Existe separador
'
'            iNumerico = IsNumeric(Left(objCombo.Text, iPosicao - 1))
'
'            If iNumerico Then
'                dNumero = Left(objCombo.Text, iPosicao - 1)
'            End If
'
'        End If
'
'    End If
'
'    iEncontrado = 0
'    If iNumerico Then   'é NÚMERO
'
'        'Verifica se é número INTEIRO
'        If dNumero <> Int(dNumero) Or dNumero > MAXIMO_INTEIRO Then Error 6729
'        'Verifica se é número NAO NEGATIVO
'        If dNumero < 0 Then Error 6742
'
'        'dNumero é inteiro positivo
'        iCodigo = dNumero
'
'        For iIndice = 0 To objCombo.ListCount - 1
'
'           'Verifica se código existe na List da ComboBox
'           If objCombo.ItemData(iIndice) = iCodigo Then
'
'                'Seleciona o ítem da List
'                objCombo.ListIndex = iIndice
'                iEncontrado = 1
'                Exit For
'
'            End If
'
'        Next
'
'        If iEncontrado = 0 Then Error 6730
'
'    Else   'é STRING
'
'        sString = Trim(objCombo.Text)
'
'        'Verifica se a string está na List da Combo
'        For iIndice = 0 To objCombo.ListCount - 1
'
'            If Mid(objCombo.List(iIndice), 1 + InStr(objCombo.List(iIndice), SEPARADOR)) Like sString & "*" Then
'
'                'Seleciona o ítem da List
'                objCombo.ListIndex = iIndice
'                iCodigo = objCombo.ItemData(iIndice)
'                iEncontrado = 1
'                Exit For
'
'            End If
'
'        Next
'
'        If iEncontrado = 0 Then Error 6731
'
'
'    End If
'
'    Combo_Seleciona = SUCESSO
'
'    Exit Function
'
'Erro_Combo_Seleciona:
'
'    Combo_Seleciona = Err
'
'    Select Case Err
'
'        Case 6729
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INTEIRO", Err, CStr(dNumero))
'            objCombo.SetFocus
'
'        Case 6742
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NEGATIVO", Err, CStr(dNumero))
'            objCombo.SetFocus
'
'        Case 6730   'Código não encontrado. Será tratado na tela chamadora
'        'Deverá ir até o BD pesquisar o objeto pelo código
'        'Se encontrar coloca na tela, se não dá aviso
'
'        Case 6731   'String não encontrada. Será tratado na tela chamadora
'        'Deverá dar mensagem de erro: objeto com nome/descrição <nome> não encontrado
'        'Segura o foco
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142326)
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Seleciona o ítem da Combo correspondente ao Text que pode ser numérico-long ou string
'Function LCombo_Seleciona(objComboBox As Object, lCodigo As Long) As Long
'
'Dim objCombo As ComboBox
'Dim iNumerico As Integer
'Dim iEncontrado As Integer
'Dim lErro As Long
'Dim dNumero As Double
'Dim iIndice As Integer
'Dim sString As String
'Dim iPosicao As String
'
'On Error GoTo Erro_LCombo_Seleciona
'
'    Set objCombo = objComboBox
'
'    'Verifica se é o código numérico
'    iNumerico = IsNumeric(objCombo.Text)
'
'    If iNumerico Then
'
'        dNumero = CDbl(objCombo.Text)
'
'    Else  'Não é número puro
'        'Determina existência e posição do caracter separador
'        iPosicao = InStr(objCombo.Text, SEPARADOR)
'
'        If iPosicao = 0 Then    'Não existe separador
'
'            iNumerico = False
'
'        Else                    'Existe separador
'
'            iNumerico = IsNumeric(Left(objCombo.Text, iPosicao - 1))
'
'            If iNumerico Then
'                dNumero = Left(objCombo.Text, iPosicao - 1)
'            End If
'
'        End If
'
'    End If
'
'    iEncontrado = 0
'    If iNumerico Then   'é NÚMERO
'
'        'Verifica se é número LONG
'        If dNumero <> Int(dNumero) Or dNumero > MAXIMO_LONG Then Error 25637
'
'        'Verifica se é número NAO NEGATIVO
'        If dNumero < 0 Then Error 25638
'
'        'dNumero é inteiro positivo
'        lCodigo = dNumero
'
'        For iIndice = 0 To objCombo.ListCount - 1
'
'           'Verifica se código existe na List da ComboBox
'           If objCombo.ItemData(iIndice) = lCodigo Then
'
'                'Seleciona o ítem da List
'                objCombo.ListIndex = iIndice
'                iEncontrado = 1
'                Exit For
'
'            End If
'
'        Next
'
'        If iEncontrado = 0 Then Error 25639
'
'    Else   'é STRING
'
'        sString = Trim(objCombo.Text)
'
'        'Verifica se a string está na List da Combo
'        For iIndice = 0 To objCombo.ListCount - 1
'
'            If Mid(objCombo.List(iIndice), 1 + InStr(objCombo.List(iIndice), SEPARADOR)) Like sString & "*" Then
'
'                'Seleciona o ítem da List
'                objCombo.ListIndex = iIndice
'                lCodigo = objCombo.ItemData(iIndice)
'                iEncontrado = 1
'                Exit For
'
'            End If
'
'        Next
'
'        If iEncontrado = 0 Then Error 25640
'
'
'    End If
'
'    LCombo_Seleciona = SUCESSO
'
'    Exit Function
'
'Erro_LCombo_Seleciona:
'
'    LCombo_Seleciona = Err
'
'    Select Case Err
'
'        Case 25637
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_LONG", Err, CStr(dNumero))
'            objCombo.SetFocus
'
'        Case 25638
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NEGATIVO", Err, CStr(dNumero))
'            objCombo.SetFocus
'
'        Case 25639   'Código não encontrado. Será tratado na tela chamadora
'        'Deverá ir até o BD pesquisar o objeto pelo código
'        'Se encontrar coloca na tela, se não dá aviso
'
'        Case 25640   'String não encontrada. Será tratado na tela chamadora
'        'Deverá dar mensagem de erro: objeto com nome/descrição <nome> não encontrado
'        'Segura o foco
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142327)
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Seleciona o ítem da ListBox correspondente a sTextBox (MaskedEdit ou TextBox) que pode ser numérico ou string
'Function List_Seleciona(objMaskedEdit As Object, objListBox As Object, iCodigo As Integer) As Long
'
'Dim objList As ListBox
'Dim iNumerico As Integer
'Dim iEncontrado As Integer
'Dim lErro As Long
'Dim dNumero As Double
'Dim iIndice As Integer
'Dim sString As String
'Dim iPosicao As String
'
'On Error GoTo Erro_List_Seleciona
'
'    Set objList = objListBox
'
'    'Testa se objeto é tipo MaskedEdit ou ListBox
'    If Not (TypeName(objMaskedEdit)= "MaskEdBox" Or TypeName(objMaskedEdit)= "TextBox") Then Error 25632
'
'    'Verifica se é o código numérico
'    iNumerico = IsNumeric(objMaskedEdit.Text)
'
'    If iNumerico Then
'
'        dNumero = CDbl(objMaskedEdit.Text)
'
'    Else  'Não é número puro
'        'Determina existência e posição do caracter separador
'        iPosicao = InStr(objMaskedEdit.Text, SEPARADOR)
'
'        If iPosicao = 0 Then    'Não existe separador
'
'            iNumerico = False
'
'        Else                    'Existe separador
'
'            iNumerico = IsNumeric(Left(objMaskedEdit.Text, iPosicao - 1))
'
'            If iNumerico Then
'                dNumero = Left(objMaskedEdit.Text, iPosicao - 1)
'            End If
'
'        End If
'
'    End If
'
'    iEncontrado = 0
'    If iNumerico Then   'é NÚMERO
'
'        'Verifica se é número INTEIRO
'        If dNumero <> Int(dNumero) Or dNumero > MAXIMO_INTEIRO Then Error 25633
'        'Verifica se é número NAO NEGATIVO
'        If dNumero < 0 Then Error 25634
'
'        'dNumero é inteiro positivo
'        iCodigo = dNumero
'
'        For iIndice = 0 To objList.ListCount - 1
'
'           'Verifica se código existe na List da ComboBox
'           If objList.ItemData(iIndice) = iCodigo Then
'
'                'Seleciona o ítem da List
'                objList.ListIndex = iIndice
'                iEncontrado = 1
'                Exit For
'
'            End If
'
'        Next
'
'        If iEncontrado = 0 Then Error 25635
'
'    Else   'é STRING
'
'        sString = Trim(objMaskedEdit.Text)
'
'        'Verifica se a string está na List da Combo
'        For iIndice = 0 To objList.ListCount - 1
'
'            If Mid(objList.List(iIndice), 1 + InStr(objList.List(iIndice), SEPARADOR)) Like sString & "*" Then
'
'                'Seleciona o ítem da List
'                objList.ListIndex = iIndice
'                iCodigo = objList.ItemData(iIndice)
'                iEncontrado = 1
'                Exit For
'
'            End If
'
'        Next
'
'        If iEncontrado = 0 Then Error 25636
'
'
'    End If
'
'    List_Seleciona = SUCESSO
'
'    Exit Function
'
'Erro_List_Seleciona:
'
'    List_Seleciona = Err
'
'    Select Case Err
'
'        Case 25632
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_OBJETO_NAO_TEM_TIPO_ESPERADO", Err, objMaskedEdit.Name)
'
'        Case 25633
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INTEIRO", Err)
'            objMaskedEdit.SetFocus
'
'        Case 25634
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NEGATIVO", Err)
'            objMaskedEdit.SetFocus
'
'        Case 25635   'Código não encontrado. Será tratado na tela chamadora
'        'Deverá ir até o BD pesquisar o objeto pelo código
'        'Se encontrar coloca na tela, se não dá aviso
'
'        Case 25636   'String não encontrada. Será tratado na tela chamadora
'        'Deverá dar mensagem de erro: objeto com nome/descrição <nome> não encontrado
'        'Segura o foco
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142328)
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Seleciona o ítem da Combo correspondente ao Text que pode ser numérico ou string
''Para ser usada em Grids no lugar de Combo_Seleciona.
''Não segura foco em caso de erro.
'Function Combo_Seleciona_Grid(objComboBox As Object, iCodigo As Integer) As Long
'
'Dim objCombo As ComboBox
'Dim iNumerico As Integer
'Dim iEncontrado As Integer
'Dim lErro As Long
'Dim dNumero As Double
'Dim iIndice As Integer
'Dim sString As String
'Dim iPosicao As String
'
'On Error GoTo Erro_Combo_Seleciona_Grid
'
'    Set objCombo = objComboBox
'
'    'Verifica se é o código numérico
'    iNumerico = IsNumeric(objCombo.Text)
'
'    If iNumerico Then
'
'        dNumero = CDbl(objCombo.Text)
'
'    Else  'Não é número puro
'        'Determina existência e posição do caracter separador
'        iPosicao = InStr(objCombo.Text, SEPARADOR)
'
'        If iPosicao = 0 Then    'Não existe separador
'
'            iNumerico = False
'
'        Else                    'Existe separador
'
'            iNumerico = IsNumeric(Left(objCombo.Text, iPosicao - 1))
'
'            If iNumerico Then
'                dNumero = Left(objCombo.Text, iPosicao - 1)
'            End If
'
'        End If
'
'    End If
'
'    iEncontrado = 0
'    If iNumerico Then   'é NÚMERO
'
'        'Verifica se é número INTEIRO
'        If dNumero <> Int(dNumero) Or dNumero > MAXIMO_INTEIRO Then Error 25083
'        'Verifica se é número POSITIVO
'        If dNumero < 1 Then Error 25084
'
'        'dNumero é inteiro positivo
'        iCodigo = dNumero
'
'        For iIndice = 0 To objCombo.ListCount - 1
'
'           'Verifica se código existe na List da ComboBox
'           If objCombo.ItemData(iIndice) = iCodigo Then
'
'                'Seleciona o ítem da List
'                objCombo.ListIndex = iIndice
'                iEncontrado = 1
'                Exit For
'
'            End If
'
'        Next
'
'        If iEncontrado = 0 Then Error 25085
'
'    Else   'é STRING
'
'        sString = Trim(objCombo.Text)
'
'        'Verifica se a string está na List da Combo
'        For iIndice = 0 To objCombo.ListCount - 1
'
'            If objCombo.List(iIndice) = sString Or Mid(objCombo.List(iIndice), 1 + InStr(objCombo.List(iIndice), SEPARADOR)) Like sString & "*" Then
'
'                'Seleciona o ítem da List
'                objCombo.ListIndex = iIndice
'                iCodigo = objCombo.ItemData(iIndice)
'                iEncontrado = 1
'                Exit For
'
'            End If
'
'        Next
'
'        If iEncontrado = 0 Then Error 25086
'
'
'    End If
'
'    Combo_Seleciona_Grid = SUCESSO
'
'    Exit Function
'
'Erro_Combo_Seleciona_Grid:
'
'    Combo_Seleciona_Grid = Err
'
'    Select Case Err
'
'        Case 25083
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INTEIRO", Err)
'
'        Case 25084
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err)
'
'        Case 25085   'Código não encontrado. Será tratado na tela chamadora
'        'Deverá ir até o BD pesquisar o objeto pelo código
'        'Se encontrar coloca na tela, se não dá aviso
'
'        Case 25086   'String não encontrada. Será tratado na tela chamadora
'        'Deverá dar mensagem de erro: objeto com nome/descrição <nome> não encontrado
'        'Segura o foco
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142329)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Combo_Item_Seleciona(objComboBox As Object) As Long
''Seleciona ítem na ComboBox referente ao codigo ( = 4 primeiros caracteres de objComboBox.text) passado por parametro
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim iCodigo As Integer
'Dim iSucesso As Integer
'
'On Error GoTo Error_Combo_Item_Seleciona
'
'    iSucesso = 0
'
'    iCodigo = Codigo_Extrai(Trim(objComboBox.Text))
'
'    'percorre todos os elementos da ComboBox
'    For iIndice = 0 To objComboBox.ListCount - 1
'
'        'compara se codigo ja existe na ComboBox
'        If objComboBox.ItemData(iIndice) = iCodigo Then
'
'            iSucesso = 1
'
'            'seleciona o item na ComboBox
'            objComboBox.ListIndex = iIndice
'            Exit For
'
'        End If
'
'    Next
'
'    If iSucesso <> 1 Then Error 12250
'
'    Combo_Item_Seleciona = SUCESSO
'
'Exit Function
'
'Error_Combo_Item_Seleciona:
'
'    Combo_Item_Seleciona = Err
'
'    Select Case Err
'
'        Case 12250   'Não encontrou o ítem na List da Combo
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142330)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Preenche_Espacos(lCodigo As Long, sEspacos As String) As Long
'
'Select Case lCodigo
'
'    Case Is < 10
'        sEspacos = "   "
'
'    Case Is < 100
'        sEspacos = "  "
'
'    Case Is < 1000
'        sEspacos = " "
'
'    Case Else
'        sEspacos = ""
'
'End Select
'
'End Function
'
'Function Dias_Mes(iMes As Integer, iAno As Integer) As Integer
''Retorna o ultima dia do mes/ano passado como parametro
'
'Dim iDiasMes As Integer
'
'    'número de dias que o último mês possui
'    Select Case iMes
'        Case 2
'            iDiasMes = 28
'        Case 4, 6, 9, 11
'            iDiasMes = 30
'        Case Else
'            iDiasMes = 31
'    End Select
'
'    If iMes = 2 Then
'        If ((iAno Mod 4) = 0 And (iAno Mod 100 <> 0)) Or ((iAno Mod 400) = 0) Then iDiasMes = 29
'    End If
'
'    Dias_Mes = iDiasMes
'
'End Function
'
'Function TabStrip_PodeTrocarTab(iFrameAtual As Integer, Opcao As Object, objForm As Object, Optional iTestaCTB As Integer = 1) As Long
''provoca o lost focus do controle ativo durante o evento click da TabStrip Opcao
'' para fazer com que o tratamento de erro seja feita com o controle ainda visivel
'' se isto gerar algum erro ou aviso a troca de frame é desfeita
'
'Dim lErro As Long
'Dim iFrameProximo As Integer
'
'On Error GoTo Erro_TabStrip_PodeTrocarTab
'
'    If GL_lUltimoErro <> -9999 Then Error 20760
'
'    GL_lUltimoErro = 0
'
'    'para forçar lost focus
'    Opcao.SetFocus
'    DoEvents
'
'    'se ocorreu algum erro ou aviso
'    If GL_lUltimoErro <> 0 Then
'        Set Opcao.SelectedItem = Opcao.Tabs(iFrameAtual)
'        Error 22884
'    End If
'
'    If iTestaCTB = 1 Then
'
'        If gcolModulo.Ativo(MODULO_CONTABILIDADE) <> MODULO_ATIVO Then
'
'            If Opcao.SelectedItem.Caption = TITULO_CONTABILIZACAO Then
'
'                Set Opcao.SelectedItem = Opcao.Tabs(iFrameAtual)
'                Error 19361
'
'            End If
'
'        End If
'
'    End If
'
'    GL_lUltimoErro = -9999
'
'    TabStrip_PodeTrocarTab = SUCESSO
'
'    Exit Function
'
'Erro_TabStrip_PodeTrocarTab:
'
'    TabStrip_PodeTrocarTab = Err
'
'    Select Case Err
'
'        Case 19361, 20760, 22884
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142331)
'
'    End Select
'
'    GL_lUltimoErro = -9999
'
'    Exit Function
'
'End Function


Function Produto_Formata(ByVal sProduto As String, sProdutoFormatado As String, iProdutoPreenchido As Integer, Optional ByVal bSemMsgErro As Boolean = False) As Long
'Devolve o produto no formato do BD: zeros a esquerda em segmentos numericos e brancos a direita para segmentos alfabeticos
'A String de Entrada vem com os caracteres delimitadores e brancos(se for o caso). O comprimento dela é o comprimento total da máscara.
'iProdutoPreenchido retornará PRODUTO_VAZIO se nenhum segmento do produto for preenchido,
'caso contrário retornará PRODUTO_PREENCHIDO
    
Dim lErro As Long
Dim iNumSegmentos As Integer
Dim iInd1 As Integer
Dim iInd2 As Integer
Dim objSegmento As New ClassSegmento
Dim sProdutoBruto As String
Dim sSegmentoBruto As String 'Segmento sendo formatado
Dim sSegmentoFormatado As String 'Segmento sendo formatado
Dim sCaracter As String 'caracter sendo tratado
Dim iRestoVazio As Integer
Dim iSegmentoVazio As Integer

On Error GoTo Erro_Produto_Formata
            
    iRestoVazio = SEGMENTOS_RESTANTES_CHEIOS
    iProdutoPreenchido = PRODUTO_VAZIO
    sProdutoBruto = sProduto
    sProdutoFormatado = ""
    iNumSegmentos = NumSegProduto()

    For iInd1 = 1 To iNumSegmentos
        
        lErro = SegmentoProduto(iInd1, objSegmento)
        If lErro <> SUCESSO Then Error 25035
        
        sSegmentoBruto = left(sProdutoBruto, objSegmento.iTamanho)
        sProdutoBruto = Mid(sProdutoBruto, objSegmento.iTamanho + Len(objSegmento.sDelimitador) + 1)
        sSegmentoFormatado = ""

        iSegmentoVazio = SEGMENTO_VAZIO

        Select Case objSegmento.iTipo
        
            Case SEGMENTO_NUMERICO

                For iInd2 = 1 To objSegmento.iTamanho
    
                    sCaracter = Mid(sSegmentoBruto, iInd2, 1)
    
                    If sCaracter = " " Or sCaracter = "" Then
                        sSegmentoFormatado = "0" & sSegmentoFormatado
                        
                    ElseIf sCaracter = "0" Then
                        sSegmentoFormatado = sSegmentoFormatado & sCaracter
                        
                    ElseIf sCaracter = "-" Then
                        Error 55692
                    Else
                        iSegmentoVazio = SEGMENTO_CHEIO
                        sSegmentoFormatado = sSegmentoFormatado & sCaracter
                    End If
    
                Next

            Case SEGMENTO_ALFANUMERICO, SEGMENTO_ASCII

                For iInd2 = objSegmento.iTamanho To 1 Step -1
    
                    sCaracter = Mid(sSegmentoBruto, iInd2, 1)
    
                    If sCaracter = " " Or sCaracter = "" Then
                        sSegmentoFormatado = " " & sSegmentoFormatado
                    Else
                        iSegmentoVazio = SEGMENTO_CHEIO
                        sSegmentoFormatado = UCase(sCaracter) & sSegmentoFormatado
                    End If
    
                Next

        End Select

        If iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS And iSegmentoVazio <> SEGMENTO_VAZIO Then Error 20736
            
        If iSegmentoVazio = SEGMENTO_VAZIO Then
            iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS
        Else
            iProdutoPreenchido = PRODUTO_PREENCHIDO
        End If


        sProdutoFormatado = sProdutoFormatado & sSegmentoFormatado

    Next
    
    If iProdutoPreenchido = PRODUTO_VAZIO Then sProdutoFormatado = ""
    
    Produto_Formata = SUCESSO
    
    Exit Function
    
Erro_Produto_Formata:

    Produto_Formata = Err

    Select Case Err

        Case 20736
            If Not bSemMsgErro Then Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEG_MEIO_NAO_PREENCHIDOS", Err)

        Case 25035

        Case 55692
            If Not bSemMsgErro Then Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEG_NUM_CARACTER_INVALIDO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142332)

    End Select

    Exit Function
    
End Function

Function Item_Formata(ByVal sNomeMascara As String, ByVal sItem As String, sItemFormatado As String, iItemPreenchido As Integer) As Long
'Devolve o item no formato do BD: zeros a esquerda em segmentos numericos e brancos a direita para segmentos alfabeticos
'A String de Entrada vem com os caracteres delimitadores e brancos(se for o caso). O comprimento dela é o comprimento total da máscara.
'iItem retornará ITEM_VAZIO se nenhum segmento do item for preenchido,
'caso contrário retornará ITEM_PREENCHIDO
    
Dim lErro As Long
Dim iNumSegmentos As Integer
Dim iInd1 As Integer
Dim iInd2 As Integer
Dim objSegmento As New ClassSegmento
Dim sItemBruto As String
Dim sSegmentoBruto As String 'Segmento sendo formatado
Dim sSegmentoFormatado As String 'Segmento sendo formatado
Dim sCaracter As String 'caracter sendo tratado
Dim iRestoVazio As Integer
Dim iSegmentoVazio As Integer

On Error GoTo Erro_Item_Formata
            
    iRestoVazio = SEGMENTOS_RESTANTES_CHEIOS
    iItemPreenchido = ITEM_VAZIO
    sItemBruto = sItem
    sItemFormatado = ""
    iNumSegmentos = NumSegItem(sNomeMascara)

    For iInd1 = 1 To iNumSegmentos
        
        lErro = SegmentoItem(sNomeMascara, iInd1, objSegmento)
        If lErro <> SUCESSO Then gError 105694
        
        sSegmentoBruto = left(sItemBruto, objSegmento.iTamanho)
        sItemBruto = Mid(sItemBruto, objSegmento.iTamanho + Len(objSegmento.sDelimitador) + 1)
        sSegmentoFormatado = ""

        iSegmentoVazio = SEGMENTO_VAZIO

        Select Case objSegmento.iTipo
        
            Case SEGMENTO_NUMERICO

                For iInd2 = 1 To objSegmento.iTamanho
    
                    sCaracter = Mid(sSegmentoBruto, iInd2, 1)
    
                    If sCaracter = " " Or sCaracter = "" Then
                        sSegmentoFormatado = "0" & sSegmentoFormatado
                        
                    ElseIf sCaracter = "0" Then
                        sSegmentoFormatado = sSegmentoFormatado & sCaracter
                        
                    ElseIf sCaracter = "-" Then
                        gError 105696
                    Else
                        iSegmentoVazio = SEGMENTO_CHEIO
                        sSegmentoFormatado = sSegmentoFormatado & sCaracter
                    End If
    
                Next

            Case SEGMENTO_ALFANUMERICO, SEGMENTO_ASCII

                For iInd2 = objSegmento.iTamanho To 1 Step -1
    
                    sCaracter = Mid(sSegmentoBruto, iInd2, 1)
    
                    If sCaracter = " " Or sCaracter = "" Then
                        sSegmentoFormatado = sSegmentoFormatado & " "
                    Else
                        iSegmentoVazio = SEGMENTO_CHEIO
                        sSegmentoFormatado = sCaracter & sSegmentoFormatado
                    End If
    
                Next

        End Select

        If iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS And iSegmentoVazio <> SEGMENTO_VAZIO Then gError 105695
            
        If iSegmentoVazio = SEGMENTO_VAZIO Then
            iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS
        Else
            iItemPreenchido = ITEM_PREENCHIDO
        End If


        sItemFormatado = sItemFormatado & sSegmentoFormatado

    Next
    
    If iItemPreenchido = ITEM_VAZIO Then sItemFormatado = ""
    
    Item_Formata = SUCESSO
    
    Exit Function
    
Erro_Item_Formata:

    Item_Formata = gErr

    Select Case gErr

        Case 105694

        Case 105695
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_SEG_MEIO_NAO_PREENCHIDOS", gErr)

        Case 105696
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_SEG_NUM_CARACTER_INVALIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142333)

    End Select

    Exit Function
    
End Function

'Function Limpa_Relatorio(Frm As Object) As Long
''limpa as telas de parametros de relatórios
'
'Dim iIndice As Integer
'Dim lErro As Long
'
'On Error GoTo Erro_Limpa_Relatorio
'
'    Call Limpa_Tela(Frm)
'
'    With Frm
'        For iIndice = 0 To .Controls.Count - 1
'            If TypeName(.Controls(iIndice))= "ComboBox" Then
'                .Controls(iIndice).ListIndex = -1
'            End If
'        Next
'    End With
'
'    Exit Function
'
'    Limpa_Relatorio = SUCESSO
'
'Erro_Limpa_Relatorio:
'
'    Limpa_Relatorio = Err
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142334)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelOpcoes_Testa_Combo(ComboBox As Object, sNome As String)
''verifica se sNOme está na combobox. Se não estiver insere. Usado na interface de relatórios
'
'Dim iIndice As Integer
'Dim lErro As Long
'
'On Error GoTo Erro_RelOpcoes_Testa_Combo:
'
'    For iIndice = 0 To ComboBox.ListCount - 1
'        If ComboBox.List(iIndice) = sNome Then Exit Function
'    Next
'
'    ComboBox.AddItem sNome
'
'    RelOpcoes_Testa_Combo = SUCESSO
'
'    Exit Function
'
'Erro_RelOpcoes_Testa_Combo:
'
'    RelOpcoes_Testa_Combo = Err
'
'        Select Case Err
'
'            Case Else
'                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142335)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function RelOpcoes_ComboOpcoes_Preenche(objRelatorio As AdmRelatorio, ComboOpcoes As Object, objRelOpcoes As AdmRelOpcoes, objTela As Object) As Long
'
'Dim colRelParametros As New Collection
'Dim lErro As Long, sCodRel As String, iIndice As Integer
'Dim objRelOpcoesAux As AdmRelOpcoes
'
'On Error GoTo Erro_RelOpcoes_ComboOpcoes_Preenche
'
'    sCodRel = objRelatorio.sCodRel
'
'    'le os nomes das opcoes do relatório existentes no BD
'    lErro = CF("RelOpcoes_Le_Todos",sCodRel, colRelParametros)
'    If lErro <> SUCESSO Then Error 48592
'
'    'preenche o ComboBox com os nomes das opções do relatório
'    For Each objRelOpcoesAux In colRelParametros
'        ComboOpcoes.AddItem objRelOpcoesAux.sNome
'    Next
'
'    'verifica se o nome da opção passada está no ComboBox
'    For iIndice = 0 To ComboOpcoes.ListCount - 1
'
'        If ComboOpcoes.List(iIndice) = objRelOpcoes.sNome Then
'
'            ComboOpcoes.Text = ComboOpcoes.List(iIndice)
'            Call objTela.PreencherParametrosNaTela(objRelOpcoes)
'
'            Exit For
'
'        End If
'
'    Next
'
'    RelOpcoes_ComboOpcoes_Preenche = SUCESSO
'
'    Exit Function
'
'Erro_RelOpcoes_ComboOpcoes_Preenche:
'
'    RelOpcoes_ComboOpcoes_Preenche = Err
'
'    Select Case Err
'
'        Case 48592
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142336)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Sub RelOpcoes_ComboOpcoes_Click(objRelOpcoes As AdmRelOpcoes, ComboOpcoes As Object, objTela As Object)
'
'Dim lErro As Long
'
'On Error GoTo Erro_RelOpcoes_ComboOpcoes_Click
'
'    If ComboOpcoes.ListIndex = -1 Then Exit Sub
'
'    objRelOpcoes.sNome = ComboOpcoes.Text
'
'    lErro = CF("RelOpcoes_Le",objRelOpcoes)
'    If (lErro <> SUCESSO) Then Error 47726
'
'    lErro = objTela.PreencherParametrosNaTela(objRelOpcoes)
'    If lErro <> SUCESSO Then Error 47727
'
'    Exit Sub
'
'Erro_RelOpcoes_ComboOpcoes_Click:
'
'    Select Case Err
'
'        Case 47726, 47727
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142337)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Function PrimeiroDia_Semana(dtData As Date, dtDataPrimeiroDia As Date) As Long
''Esta funcao retorna o primeiro dia da semana da data em questao
''"A DATA DA SEGUNDA-FEIRA"
'
'Dim iDiaSemana As Integer
'Dim lErro As Long
'
'On Error GoTo Erro_PrimeiroDia_Semana
'
'    iDiaSemana = Weekday(dtData)
'
'    Select Case iDiaSemana
'
'        Case vbSunday
'            dtDataPrimeiroDia = dtData + 1
'        Case vbMonday
'            dtDataPrimeiroDia = dtData
'        Case vbTuesday
'            dtDataPrimeiroDia = dtData - 1
'        Case vbWednesday
'            dtDataPrimeiroDia = dtData - 2
'        Case vbThursday
'            dtDataPrimeiroDia = dtData - 3
'        Case vbFriday
'            dtDataPrimeiroDia = dtData - 4
'        Case vbSaturday
'            dtDataPrimeiroDia = dtData - 5
'
'        End Select
'
'    PrimeiroDia_Semana = SUCESSO
'
'    Exit Function
'
'Erro_PrimeiroDia_Semana:
'
'    PrimeiroDia_Semana = Err
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142338)
'
'        End Select
'
'    Exit Function
'
'End Function
'
'Sub RelOpcoes_ComboOpcoes_Validate(ComboOpcoes As Object)
'
'Dim lErro As Long
'
'On Error GoTo Erro_RelOpcoes_ComboOpcoes_Validate
'
'    If ComboOpcoes.ListIndex <> -1 Then Exit Sub
'
'    lErro = Combo_Item_Igual(ComboOpcoes)
'    If lErro <> SUCESSO Then Error 56733
'
'    Exit Sub
'
'Erro_RelOpcoes_ComboOpcoes_Validate:
'
'    Select Case Err
'
'        Case 56733
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142339)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Sub MaskEdBox_TrataGotFocus(objMasked As Object, Optional viAlterado As Variant)
'Dim sMascaraAntiga As String, iAlteradoAntigo As Integer
'
'    If objMasked.Mask <> "" Then
'
'        If InStr(objMasked.Mask, "#") <> 0 Or InStr(objMasked.Mask, "9") <> 0 Then
'
'            If Len(objMasked.ClipText) = 0 Then
'
'                If Not IsMissing(viAlterado) Then
'                    iAlteradoAntigo = viAlterado
'                End If
'
'                sMascaraAntiga = objMasked.Mask
'                objMasked.Mask = ""
'                objMasked.Text = ""
'                objMasked.Mask = sMascaraAntiga
'
'                If Not IsMissing(viAlterado) Then
'                    viAlterado = iAlteradoAntigo
'                End If
'
'            End If
'
'        End If
'
'    End If
'
'End Sub
'
Function SCombo_Seleciona(ByVal objComboBox As Object) As Long
'Seleciona o ítem da Combo correspondente ao Text, procura dos dois lados do Separador
'Retorna o Codigo do item Selecionado

Dim objCombo As ComboBox
Dim sTexto As String
Dim iPosicaoSeparador As Integer
Dim iEncontrado As Integer
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_SCombo_Seleciona

    Set objCombo = objComboBox

    sTexto = objCombo.Text

    For iIndice = 0 To objCombo.ListCount - 1

        iPosicaoSeparador = InStr(objCombo.List(iIndice), SEPARADOR)

        If Trim(sTexto) = Trim(Mid(objCombo.List(iIndice), 1, iPosicaoSeparador - 1)) Then

            'Seleciona o ítem da List
            objCombo.ListIndex = iIndice
            iEncontrado = 1
            Exit For
        End If
    Next

    If iEncontrado = 0 Then

        For iIndice = 0 To objCombo.ListCount - 1

            iPosicaoSeparador = InStr(objCombo.List(iIndice), SEPARADOR)

            If Trim(sTexto) = Trim(Mid(objCombo.List(iIndice), iPosicaoSeparador + 1)) Then
                'Seleciona o ítem da List
                objCombo.ListIndex = iIndice
                iEncontrado = 1
                Exit For
            End If

        Next

    End If

    If iEncontrado = 0 Then Error 60483

    SCombo_Seleciona = SUCESSO

    Exit Function

Erro_SCombo_Seleciona:

    SCombo_Seleciona = Err

    Select Case Err

        Case 60483   'String não encontrada. Será tratado na tela chamadora
        'Deverá dar mensagem de erro: objeto com nome/descrição <nome> não encontrado
        'Segura o foco

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142340)

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

'************************************** Função alterada por Luiz Gustavo de Freitas Nogueira em 30/05/2001 **************************
'Alterações: inclusão de chamada para geração de cabeçalho de gráfico, caso seja recebida uma coleção com dados de cabeçalho / rodapé
'Inclusão para tratamento de exibição de legenda
'Inclusão de tratamento de exibição de DataLabels para cada série
'************************************************************************************************************************************

Dim iColuna As Integer
Dim iLinha As Integer
Dim iEixoXPreenchido As Integer
Dim iImprimeGrafico As Integer
Dim iSerie As Integer
Dim lErro As Long
Dim sCelulaAtual As String
'Dim sCelulaInicioY As String
'Dim sCelulaInicioX As String
'Dim sCelulaFimY As String
'Dim sCelulaFimX As String
Dim lNumRegs As Long, lColunaX As Long
Dim sCabecalhoRodape As String
Dim asCabecalhoRodape(1 To 6) As String
'Dim avValoresCelulas() As Variant
Dim objColuna As ClassColunasExcel
Dim objCelulas As ClassCelulasExcel
Dim vbMsgRes As VbMsgBoxResult

'Declarações para manipulação do Excel
'Dim objPastaTrabalho As Object 'Excel.Workbook
'Dim objPlanilhaExcel As Object 'Excel.Worksheet
'Dim objGrafico As Object  'Excel.Chart
'Dim objFaixaCelulasX As Object 'Excel.Range
'Dim objFaixaCelulasY As Object 'Excel.Range
'Dim objFaixaCelulasPreencher As Object 'Excel.Range
Dim iLin As Integer
Dim iCol As Integer
Dim objExcelApp As New ClassExcelApp

On Error GoTo Erro_Excel_Cria_Grafico
    
'    'Luiz
'    'Se não existe uma instância do excel em aberto
'    If GL_objExcelSist Is Nothing Then
'
'        'Cria uma nova instância do Excel
'        Set GL_objExcelSist = CreateObject("Excel.Application")
'
'    End If
'
'    'Desabilita as mensagens do Excel
'    GL_objExcelSist.DisplayAlerts = False

    'Abre o excel
    lErro = objExcelApp.Abrir()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Pergunta se o usuário deseja exibir o gráfico na tela
    vbMsgRes = MsgBox("Deseja exibir o gráfico na tela?", vbYesNoCancel, objPlanilha.sNomePlanilha)
    If vbMsgRes = vbNo Then iImprimeGrafico = EXCEL_MODO_IMPRESSÃO_GRAFICO
    If vbMsgRes = vbCancel Then Exit Function
    
'    'Inicia uma nova Pasta de Trabalho e selciona uma planilha
'    Set objPastaTrabalho = GL_objExcelSist.Workbooks.Add
'    Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
'
'    'Redimensiona o array de acordo a quantidade de linhas e colunas recebidas
'    ReDim avValoresCelulas(1 To objPlanilha.colColunas(1).colCelulas.Count, 1 To objPlanilha.colColunas.Count)
    
    lErro = objExcelApp.Criar_Planilha()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Para cada coluna que deverá receber dados
    For iColuna = 1 To objPlanilha.colColunas.Count
        
        'Instancia objColuna como um obj da Coleção de colunas
        Set objColuna = objPlanilha.colColunas(iColuna)
        
        'E para cada linha que receberá dados
        For iLinha = 1 To objColuna.colCelulas.Count
            
            'Instancia objCelulas como um obj da Coleção de células
            Set objCelulas = objColuna.colCelulas(iLinha)
            
'            'Monta o array que será usado para preencher a planilha
'            '#########################################################
'            'Alterado por Wagner
            If IsNumeric(objCelulas.vValor) And objCelulas.vValor <> Empty Then
'                avValoresCelulas(iLinha, iColuna) = StrParaDbl(objCelulas.vValor)
'            Else
'                avValoresCelulas(iLinha, iColuna) = objCelulas.vValor
                objCelulas.vValor = StrParaDbl(objCelulas.vValor)
                objCelulas.sNumberFormat = "#,##0.00"
            End If
'            '#########################################################
            
            lErro = objExcelApp.Preencher_Celula(iLinha, iColuna, objCelulas.vValor, objCelulas.sNumberFormat, objCelulas.bFonteNegrito)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        Next
        
        lNumRegs = objColuna.colCelulas.Count
         
        'Seleciona as colunas que participarão do gráfico
        Select Case objColuna.iParticipaGrafico
            
            'Se a coluna faz parte do eixo X
            Case EXCEL_PARTICIPA_GRAFICO_X
            
                'Se já preencheu a faixa de células que formarão o eixo X => erro
                If iEixoXPreenchido = 1 Then gError 79947
                
'                'Guarda as células de início e fim da coluna atual
'                sCelulaInicioX = Chr$(64 + iColuna) & 2
'                sCelulaFimX = Chr$(64 + iColuna) & objColuna.colCelulas.Count

'                'Monta a Faixa de Células que participarão do Eixo X
'                Set objFaixaCelulasX = objPlanilhaExcel.Range(sCelulaInicioX, sCelulaFimX)

                lColunaX = iColuna
                
                'Informa que já foram preenchidas as informações para o eixo X
                iEixoXPreenchido = 1
                
'                If objColuna.dLarguraColuna > 0 Then
'                    objFaixaCelulasX.ColumnWidth = objColuna.dLarguraColuna
'                End If
        
            'Se a coluna faz parte do eixo Y
            Case EXCEL_PARTICIPA_GRAFICO_Y
                
'                'Guarda as células de início e fim da coluna atual
'                sCelulaInicioY = Chr$(64 + iColuna) & 2
'                sCelulaFimY = Chr$(64 + iColuna) & objColuna.colCelulas.Count
                
'                If objFaixaCelulasY Is Nothing Then
'                    Set objFaixaCelulasY = objPlanilhaExcel.Range(sCelulaInicioY, sCelulaFimY)
'                Else
'                    Set objFaixaCelulasY = GL_objExcelSist.Union(objFaixaCelulasY, objPlanilhaExcel.Range(sCelulaInicioY, sCelulaFimY))
'                End If
'
'                If objColuna.dLarguraColuna > 0 Then
'                    objFaixaCelulasY.ColumnWidth = objColuna.dLarguraColuna
'                End If
            
            Case Else
'                objPlanilhaExcel.Range(Chr$(64 + iColuna) & 2, Chr$(64 + iColuna) & objColuna.colCelulas.Count).ColumnWidth = objColuna.dLarguraColuna
            
        End Select
        
    Next
    
    lErro = objExcelApp.Ajustar_Tam_Colunas()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
'    'Monta a Faixa das Células que receberão os valores passado
'    Set objFaixaCelulasPreencher = objPlanilhaExcel.Range("A1", Chr$(64 + objPlanilha.colColunas.Count) & CStr(objPlanilha.colColunas(1).colCelulas.Count))
'
'    'Passa os valores armazenados no array para a faixa determinada acima
'    objFaixaCelulasPreencher.Value = avValoresCelulas
'
'    '########################################################
'    'Inserido por Wagner 01/03/2007
'    For iCol = 1 To objPlanilha.colColunas.Count
'
'        For iLin = 1 To objPlanilha.colColunas(iCol).colCelulas.Count
'
'            If objPlanilha.colColunas(iCol).colCelulas.Item(iLin).bFonteNegrito Then
'                objFaixaCelulasPreencher(iLin, iCol).Font.Bold = True
'            End If
'
'            If Len(Trim(objPlanilha.colColunas(iCol).colCelulas.Item(iLin).sNumberFormat)) > 0 Then
'                objFaixaCelulasPreencher(iLin, iCol).NumberFormat = objPlanilha.colColunas(iCol).colCelulas.Item(iLin).sNumberFormat
'            End If
'        Next
'
'    Next
'    '########################################################
    
'    'Adiciona um gráfico à planilha
'    Set objGrafico = objPastaTrabalho.Charts.Add

    lErro = objExcelApp.Criar_Grafico(1, lNumRegs, 1, objPlanilha.colColunas.Count, objPlanilha.sTituloGrafico, objPlanilha.vPlotLinhaColuna, objPlanilha.lPosicaoLegenda, objPlanilha.iEixosGrafico, objPlanilha.lLabelsXPosicao, objPlanilha.lLabelsXOrientacao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Se foi passado pelo menos uma linha na coleção de linhas do cabeçalho
    If Not (objPlanilha.colCabecalhoRodape Is Nothing) Then
    
        'Monta as seções de cabeçalho e rodapé do Gráfico
        lErro = Excel_Obtem_Cabecalho(objPlanilha.colCabecalhoRodape, asCabecalhoRodape)
        If lErro <> SUCESSO Then gError 90534
    
        'Preenche as seções de cabeçalho
'        If Len(Trim(asCabecalhoRodape(EXCEL_CABECALHO_ESQUERDO))) > 0 Then objGrafico.PageSetup.LeftHeader = asCabecalhoRodape(EXCEL_CABECALHO_ESQUERDO)
'        If Len(Trim(asCabecalhoRodape(EXCEL_CABECALHO_CENTRAL))) > 0 Then objGrafico.PageSetup.CenterHeader = asCabecalhoRodape(EXCEL_CABECALHO_CENTRAL)
'        If Len(Trim(asCabecalhoRodape(EXCEL_CABECALHO_DIREITO))) > 0 Then objGrafico.PageSetup.RightHeader = asCabecalhoRodape(EXCEL_CABECALHO_DIREITO)
'        If Len(Trim(asCabecalhoRodape(EXCEL_RODAPE_ESQUERDO))) > 0 Then objGrafico.PageSetup.LeftFooter = asCabecalhoRodape(EXCEL_RODAPE_ESQUERDO)
'        If Len(Trim(asCabecalhoRodape(EXCEL_RODAPE_CENTRAL))) > 0 Then objGrafico.PageSetup.CenterFooter = asCabecalhoRodape(EXCEL_RODAPE_CENTRAL)
'        If Len(Trim(asCabecalhoRodape(EXCEL_RODAPE_DIREITO))) > 0 Then objGrafico.PageSetup.RightFooter = asCabecalhoRodape(EXCEL_RODAPE_DIREITO)
    
        If Len(Trim(asCabecalhoRodape(EXCEL_CABECALHO_ESQUERDO))) > 0 Then Call objExcelApp.Grafico_Set_LeftHeader(asCabecalhoRodape(EXCEL_CABECALHO_ESQUERDO))
        If Len(Trim(asCabecalhoRodape(EXCEL_CABECALHO_CENTRAL))) > 0 Then Call objExcelApp.Grafico_Set_CenterHeader(asCabecalhoRodape(EXCEL_CABECALHO_CENTRAL))
        If Len(Trim(asCabecalhoRodape(EXCEL_CABECALHO_DIREITO))) > 0 Then Call objExcelApp.Grafico_Set_RightHeader(asCabecalhoRodape(EXCEL_CABECALHO_DIREITO))
        If Len(Trim(asCabecalhoRodape(EXCEL_RODAPE_ESQUERDO))) > 0 Then Call objExcelApp.Grafico_Set_LeftFooter(asCabecalhoRodape(EXCEL_RODAPE_ESQUERDO))
        If Len(Trim(asCabecalhoRodape(EXCEL_RODAPE_CENTRAL))) > 0 Then Call objExcelApp.Grafico_Set_CenterFooter(asCabecalhoRodape(EXCEL_RODAPE_CENTRAL))
        If Len(Trim(asCabecalhoRodape(EXCEL_RODAPE_DIREITO))) > 0 Then Call objExcelApp.Grafico_Set_RightFooter(asCabecalhoRodape(EXCEL_RODAPE_DIREITO))
    
    End If

'    With objGrafico
'
'        'Monta o Gráfico utilizando a Faixa de células que formarão o eixo Y; informa se as séries do gráfico serão por linha ou coluna e informa o título do gráfico
'        .ChartWizard objFaixaCelulasY, , , objPlanilha.vPlotLinhaColuna, , , , objPlanilha.sTituloGrafico
'
'        'Determina os valores do Eixo X
'        .SeriesCollection(1).XValues = objFaixaCelulasX
'
'        'Se o gráfico foi setado para não exibir legenda
'        If objPlanilha.lPosicaoLegenda = EXCEL_LEGENDA_NAO_EXIBE Then
'
'            'Informa ao excel para ocultar a legenda
'            .HasLegend = False
'
'        'Se foi setado para exibir a legenda
'        Else
'
'            'Informa ao excel para exibir a legenda
'            .HasLegend = True
'
'            'Posiciona a legenda do gráfico
'            .Legend.Position = objPlanilha.lPosicaoLegenda
'
'        End If
        
        'Para cada coluna existente na planilha
        For iColuna = 1 To objPlanilha.colColunas.Count
            
            'Instancia objColuna com os valores da coluna que está sendo percorrida
            Set objColuna = objPlanilha.colColunas(iColuna)
            
            'Se a coluna participa do gráfico no eixo Y
            If objColuna.iParticipaGrafico = EXCEL_PARTICIPA_GRAFICO_Y Then
            
'                'Determina o número da série de dados
                iSerie = iSerie + 1
'
'                'Instancia objCelulas como a primeira célula da coluna
'                Set objCelulas = objColuna.colCelulas(1)
'
'                'Define o nome da série de dados
'                .SeriesCollection(iSerie).Name = "=" & Chr$(34) & objCelulas.vValor & Chr$(34)
'
'                'Define o tipo do gráfico da série de dados
'                .SeriesCollection(iSerie).ChartType = objColuna.lTipoGraficoColuna
'
'                'Define o que será exibido como label da série (Valor, Legenda, Percentual, Legenda e Percentual, Não exibe)
'                .SeriesCollection(iSerie).ApplyDataLabels (objColuna.lDataLabels)
'
'                'Se os labels de valores do eixo X forem exibidos => define a orientação dos labels
'                If objColuna.lDataLabels <> EXCEL_NAO_EXIBE_LABELS Then .SeriesCollection(iSerie).DataLabels.Orientation = objColuna.lDataLabelsOrientacao

                lErro = objExcelApp.Incluir_Campo_Grafico(lNumRegs, iColuna, lColunaX, iSerie, objColuna.lTipoGraficoColuna, objColuna.lDataLabels, objColuna.lDataLabelsOrientacao)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
            End If
        
        Next
    
'        'Se o gráfico, é um gráfico que possui eixos
'        If objPlanilha.iEixosGrafico = EXCEL_GRAFICO_COM_EIXOS Then
'
'            'Determina a posição dos labels do eixo X
'            .Axes(EXCEL_XLCATEGORY).TickLabelPosition = objPlanilha.lLabelsXPosicao
'
'            'Determina a orientação dos labels do eixo X
'            .Axes(EXCEL_XLCATEGORY).TickLabels.Orientation = objPlanilha.lLabelsXOrientacao
'        End If
'
'    End With
'
'    'Define os nomes da planilha e do gráfico
'    objPlanilhaExcel.Name = objPlanilha.sNomePlanilha
'    objGrafico.Name = objPlanilha.sNomeGrafico
    
    lErro = objExcelApp.Muda_Nome_Aba(objPlanilha.sNomePlanilha)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = objExcelApp.Muda_Nome_Grafico(objPlanilha.sNomeGrafico)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If iImprimeGrafico = EXCEL_MODO_IMPRESSÃO_GRAFICO Then
        'Imprime o gráfico e fecha o excel
'        objGrafico.PrintOut

        lErro = objExcelApp.Imprimir_Grafico()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'        GL_objExcelSist.Quit

        lErro = objExcelApp.Fechar()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Else
'        'Torna o Excel Visível
'        GL_objExcelSist.Visible = True

        lErro = objExcelApp.Mudar_Visibilidade(True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Excel_Cria_Grafico = SUCESSO
    
    Exit Function
        
Erro_Excel_Cria_Grafico:

    Excel_Cria_Grafico = gErr
    
    Select Case gErr
            
        Case 90534
        
        Case 79947
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCEL_EIXO_X_JA_DEFINIDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142341)
        
    End Select
    
    'Fecha o Excel
    'GL_objExcelSist.Quit
    Call objExcelApp.Fechar
    
    Exit Function
    
End Function

Public Function Excel_Obtem_Cabecalho(colCabecalhoRodape As Collection, asCabecalhoRodape() As String) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim sQuebraLinha As String
Dim iTamanhoCabecalho As Integer
Dim objLinha As ClassLinhaCabecalhoExcel

On Error GoTo Erro_Excel_Obtem_Cabecalho
    
    'Ordena as linhas passadas em colCabecalhoRodape por Seção, Linha
    lErro = Excel_Ordena_Cabecalho_Rodape(colCabecalhoRodape)
    If lErro <> SUCESSO Then gError 90533
    
    'Para cada Linha passada na coleção
    For Each objLinha In colCabecalhoRodape
        
        'Se for a primeira linha
        If objLinha.iLinha = 1 Then
        
            'Monta a string com o texto e os parâmetros de formtação do texto no excel, INLCUINDO O ALINHAMENTO
            objLinha.sTexto = Trim(objLinha.sAlinhamento) & "&" & Chr$(34) & Trim(objLinha.sFonte) & Chr$(34) & "&" & Trim(objLinha.iTamanhoFonte) & Trim(objLinha.sItalico) & Trim(objLinha.sNegrito) & Trim(objLinha.sSublinhado) & "  " & objLinha.sTexto & Trim(objLinha.sItalico) & Trim(objLinha.sNegrito) & Trim(objLinha.sSublinhado)
        
        Else
        
            'Monta a string com o texto e os parâmetros de formtação do texto no excel, MAS NÃO INCLUI O ALINHAMENTO
            objLinha.sTexto = "&" & Chr$(34) & Trim(objLinha.sFonte) & Chr$(34) & "&" & Trim(objLinha.iTamanhoFonte) & Trim(objLinha.sItalico) & Trim(objLinha.sNegrito) & Trim(objLinha.sSublinhado) & "  " & objLinha.sTexto & Trim(objLinha.sItalico) & Trim(objLinha.sNegrito) & Trim(objLinha.sSublinhado)
        
        End If
        
        'Se foi configurado para inserir linhas em branco entre essa linha e a próxima
        If objLinha.iEspacoLinha > 0 Then
            
            'Limpa a variável
            sQuebraLinha = ""
            
            'Para cada linha que deve ser inserida em branco
            For iIndice = 1 To objLinha.iEspacoLinha
            
                'Inclui uma quebra
                sQuebraLinha = sQuebraLinha & vbCr
            
            Next
        
        'Se não foi configurado nenhum espaço entre linhas
        ElseIf objLinha.iEspacoLinha <> EXCEL_CABECALHO_RODAPE_NAO_QUEBRA_LINHA Then
        
            'Inclui apenas uma quebra
            objLinha.sTexto = objLinha.sTexto & vbCr
        
        End If
            
        objLinha.sTexto = objLinha.sTexto & sQuebraLinha
        
        'Limpa a variável
        sQuebraLinha = ""
        
        Select Case objLinha.iSecao
        
            Case EXCEL_CABECALHO_ESQUERDO
                
                asCabecalhoRodape(EXCEL_CABECALHO_ESQUERDO) = asCabecalhoRodape(EXCEL_CABECALHO_ESQUERDO) & objLinha.sTexto
    
            Case EXCEL_CABECALHO_CENTRAL
                
                asCabecalhoRodape(EXCEL_CABECALHO_CENTRAL) = asCabecalhoRodape(EXCEL_CABECALHO_CENTRAL) & objLinha.sTexto
            
            Case EXCEL_CABECALHO_DIREITO
                
                asCabecalhoRodape(EXCEL_CABECALHO_DIREITO) = asCabecalhoRodape(EXCEL_CABECALHO_DIREITO) & objLinha.sTexto
            
            Case EXCEL_RODAPE_ESQUERDO
                
                asCabecalhoRodape(EXCEL_RODAPE_ESQUERDO) = asCabecalhoRodape(EXCEL_RODAPE_ESQUERDO) & objLinha.sTexto
            
            Case EXCEL_RODAPE_CENTRAL
                
                asCabecalhoRodape(EXCEL_RODAPE_CENTRAL) = asCabecalhoRodape(EXCEL_RODAPE_CENTRAL) & objLinha.sTexto
            
            Case EXCEL_RODAPE_DIREITO
                
                asCabecalhoRodape(EXCEL_RODAPE_DIREITO) = asCabecalhoRodape(EXCEL_RODAPE_DIREITO) & objLinha.sTexto
        
        End Select
    
    Next
    
    'Para cada seção do cabeçalho
    For iIndice = EXCEL_CABECALHO_ESQUERDO To EXCEL_CABECALHO_DIREITO
    
        iTamanhoCabecalho = iTamanhoCabecalho + Len(asCabecalhoRodape(iIndice))
    Next
    
    If iTamanhoCabecalho > EXCEL_CABECALHO_RODAPE_TAMANHO Then gError 90535
    
    'Zera a variável
    iTamanhoCabecalho = 0
    
    'Para cada seção do rodapé
    For iIndice = EXCEL_RODAPE_ESQUERDO To EXCEL_RODAPE_DIREITO
    
        'Soma o tamanho da string passada para a seção
        iTamanhoCabecalho = iTamanhoCabecalho + Len(asCabecalhoRodape(iIndice))
    Next
    
    'Se a string passada tem mais de 249 caracteres => erro
    If iTamanhoCabecalho > EXCEL_CABECALHO_RODAPE_TAMANHO Then gError 90536
    
    Excel_Obtem_Cabecalho = SUCESSO
    
    Exit Function

Erro_Excel_Obtem_Cabecalho:
    
    Excel_Obtem_Cabecalho = gErr
    
    Select Case gErr
    
        Case 90533
        
        Case 90535, 90536
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCEL_CABECALHO_RODAPE_TAMANHO", gErr, iTamanhoCabecalho)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142342)
    
    End Select
    
    Exit Function
    
End Function

Public Function Excel_Ordena_Cabecalho_Rodape(colLinhas As Collection) As Long

Dim iSecaoAux As Integer
Dim iLinhaAux As Integer
Dim iIndice As Integer
Dim iIndiceAux As Integer
Dim iItensCol As Integer
Dim colLinhasAux As New Collection

On Error GoTo Erro_Excel_Ordena_Cabecalho_Rodape
    
    iItensCol = colLinhas.Count
    
    Do
        iSecaoAux = colLinhas(1).iSecao
        iLinhaAux = colLinhas(1).iLinha
        iIndiceAux = 1
        
        For iIndice = 1 To colLinhas.Count
            
            If colLinhas(iIndice).iSecao <= iSecaoAux Then
                
                iSecaoAux = colLinhas(iIndice).iSecao
                
                If colLinhas(iIndice).iLinha < iLinhaAux Then
                    
                    iLinhaAux = colLinhas(iIndice).iLinha
                    iIndiceAux = iIndice
                
                End If
                
            
            End If
        
        Next
        
        colLinhasAux.Add colLinhas(iIndiceAux)
        
        colLinhas.Remove iIndiceAux
    
    Loop While colLinhasAux.Count <> iItensCol
    
    Set colLinhas = colLinhasAux
    
    Excel_Ordena_Cabecalho_Rodape = SUCESSO
    
    Exit Function

Erro_Excel_Ordena_Cabecalho_Rodape:

    Excel_Ordena_Cabecalho_Rodape = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142343)
    
    End Select
    
    Exit Function
    
End Function

'??? Luiz: Transportar para RotinasAdmFormata
'Revisado - OK Luiz 19/02/02
'Rotina criada por Daniel em 31/01/02
Public Function Gera_Nome_Arquivo(sArquivo As String, ByVal sExtensaoArq As String) As Long
'Esta função gera um nome de arquivo a partir de:
'sArquivo RECEBE Diretório(Path Completo com "\" no fim -> ex. "C:\Contab\") + "Prefixo" do nome do Arquivo
'sArquivo RETORNA o Diretório + Nome do Arquivo
'sExtensaoArq contém a Extensão do Arquivo (Ex: .txt, .bin)

Dim bGerou As Boolean
Dim lSequencialArq As Long
Dim sArqAux As String

On Error GoTo Erro_Gera_Nome_Arquivo

    bGerou = False
    lSequencialArq = 0

    Do While Not bGerou

        'Incrementa o sequencial do arquivo
        lSequencialArq = lSequencialArq + 1

        'Gera o nome do arquivo
        sArqAux = sArquivo & "_" & Format(gdtDataHoje, "DDMMYY") & "_" & FormataCpoNum(lSequencialArq, 2) & sExtensaoArq

        'Verifica se o nome criado já existe
        If Len(Trim(Dir(sArqAux))) = 0 Then bGerou = True

    Loop

    sArquivo = sArqAux

    Gera_Nome_Arquivo = SUCESSO

    Exit Function

Erro_Gera_Nome_Arquivo:

    Gera_Nome_Arquivo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142344)

    End Select

End Function

Function SCombo_Seleciona2(ByVal objComboBox As Object, ByVal sTexto As String) As Long
'Seleciona o ítem da Combo correspondente ao sTexto, procura dos dois lados do Separador

Dim objCombo As ComboBox
Dim iPosicaoSeparador As Integer
Dim iEncontrado As Integer
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_SCombo_Seleciona2

    Set objCombo = objComboBox

    For iIndice = 0 To objCombo.ListCount - 1

        iPosicaoSeparador = InStr(objCombo.List(iIndice), SEPARADOR)

        If iPosicaoSeparador = 0 Then Exit For
        
        If Trim(sTexto) = Trim(Mid(objCombo.List(iIndice), 1, iPosicaoSeparador - 1)) Then

            'Seleciona o ítem da List
            objCombo.ListIndex = iIndice
            iEncontrado = 1
            Exit For
        End If
    Next

    If iEncontrado = 0 Then

        For iIndice = 0 To objCombo.ListCount - 1

            iPosicaoSeparador = InStr(objCombo.List(iIndice), SEPARADOR)

            If Trim(sTexto) = Trim(Mid(objCombo.List(iIndice), iPosicaoSeparador + 1)) Then
                'Seleciona o ítem da List
                objCombo.ListIndex = iIndice
                iEncontrado = 1
                Exit For
            End If

        Next

    End If

    If iEncontrado = 0 Then gError 108941

    SCombo_Seleciona2 = SUCESSO

    Exit Function

Erro_SCombo_Seleciona2:

    SCombo_Seleciona2 = gErr

    Select Case gErr

        Case 108941   'String não encontrada. Será tratado na tela chamadora
        'Deverá dar mensagem de erro: objeto com nome/descrição <nome> não encontrado
        'Segura o foco

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142345)

    End Select

End Function

Function Produto_Segmento(objGridInt As AdmGrid, KeyAscii As Integer) As Long
'ao teclar enter faz com que o cursor se dirija para o inicio do proximo segmento

    Produto_Segmento = SUCESSO

End Function

'Incluído por Luiz Nogueira em 13/04/04
Public Sub Converte_CheckBox_SimNao(ByVal vCampo, sCheckBoxSimNao)
'Converte os valores 1 ou 0 da checkbox em Sim ou Não
    
    'Se for Zero recebe Não
    If vCampo = DESMARCADO Then
        sCheckBoxSimNao = "Não"
    'Se não recebe Sim
    Else
        sCheckBoxSimNao = "Sim"
    End If
    
    Exit Sub

End Sub

