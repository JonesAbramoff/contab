VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmFormataExt"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function MesNome(iMes As Integer, sMes As String) As Long
'Entra com o número do mês. Retorna-o por extenso.

Dim lErro As Long

On Error GoTo Erro_MesNome

    'Seleciona o mês que corresponde a iMes
    Select Case iMes
    
        Case 1
            sMes = "Janeiro"
            
        Case 2
            sMes = "Fevereiro"
            
        Case 3
            sMes = "Março"
            
        Case 4
            sMes = "Abril"
        
        Case 5
            sMes = "Maio"
        
        Case 6
            sMes = "Junho"
        
        Case 7
            sMes = "Julho"
        
        Case 8
            sMes = "Agosto"
        
        Case 9
            sMes = "Setembro"
        
        Case 10
            sMes = "Outubro"
        
        Case 11
            sMes = "Novembro"
        
        Case 12
            sMes = "Dezembro"
            
        Case Else
            Error 25208
        
    End Select
    
    MesNome = SUCESSO

    Exit Function

Erro_MesNome:

    MesNome = Err
    
    Select Case Err

        Case 25208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_NAO_MES", Err, iMes)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142346)

    End Select

    Exit Function

End Function
Function MesNumero(sMes As String, iMes As Integer) As Long
'Entra com o mês por extenso. Retorna-o como um inteiro.

Dim lErro As Long

On Error GoTo Erro_MesNumero

    'Seleciona o mês que corresponde a sMes
    Select Case sMes
    
        Case "Janeiro"
            iMes = 1
            
        Case "Fevereiro"
            iMes = 2
            
        Case "Março"
            iMes = 3
            
        Case "Abril"
            iMes = 4
        
        Case "Maio"
            iMes = 5
        
        Case "Junho"
            iMes = 6
        
        Case "Julho"
            iMes = 7
        
        Case "Agosto"
            iMes = 8
        
        Case "Setembro"
            iMes = 9
        
        Case "Outubro"
            iMes = 10
        
        Case "Novembro"
            iMes = 11
        
        Case "Dezembro"
            iMes = 12
            
        Case Else
            Error 25209
        
    End Select
    
    MesNumero = SUCESSO

    Exit Function

Erro_MesNumero:

    MesNumero = Err
    
    Select Case Err

        Case 25209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_STRING_NAO_MES", Err, sMes)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142347)

    End Select

    Exit Function

End Function

Function StrParaDate(sTexto As String) As Date
'retorna texto como data

    If Len(sTexto) = 0 Or left(sTexto, 2) = "  " Then
        StrParaDate = DATA_NULA
    Else
        StrParaDate = CDate(sTexto)
    End If
    
End Function

Function StrParaDbl(sTexto As String) As Double
'retorna sTexto como double

    If Len(Trim(sTexto)) = 0 Then
        StrParaDbl = 0
    Else
        StrParaDbl = CDbl(sTexto)
    End If

End Function


Function PercentParaDbl(sTexto As String) As Double
'retorna percentagem sTexto como double
'sTexto deve ser número finalizado por %

    If Len(Trim(sTexto)) = 0 Then
        PercentParaDbl = 0
    ElseIf Not IsNumeric(left(sTexto, Len(sTexto) - 1)) Then
        PercentParaDbl = 0
    ElseIf right(sTexto, 1) <> "%" Then
        PercentParaDbl = 0
    Else
        PercentParaDbl = CDbl(Format(sTexto, "General Number"))
    End If

End Function

Function StrParaInt(sTexto As String) As Integer
'retorna sTexto como Integer

    If Len(Trim(sTexto)) = 0 Then
        StrParaInt = 0
    Else
        StrParaInt = CInt(sTexto)
    End If

End Function

Function StrParaLong(sTexto As String) As Long
'retorna sTexto como Long

    If Len(Trim(sTexto)) = 0 Then
        StrParaLong = 0
    Else
        StrParaLong = CLng(sTexto)
    End If

End Function

Sub DateParaMasked(objMasked As Object, dtData As Date)
'preenche MaskedEdit com Date

    If dtData <> DATA_NULA Then
        objMasked.PromptInclude = False
        If objMasked.Mask = "##/##/####" Then
            objMasked.Text = Format(dtData, "dd/mm/yyyy")
        Else
            objMasked.Text = Format(dtData, "dd/mm/yy")
        End If
        objMasked.PromptInclude = True
    Else
        objMasked.PromptInclude = False
        objMasked.Text = ""
        objMasked.PromptInclude = True
    End If
        
End Sub

Function MaskedParaDate(objMasked As Object) As Date
'retorna objMasked.Text como Date

    If Len(Trim(objMasked.ClipText)) = 0 Then
        MaskedParaDate = DATA_NULA
    Else
        MaskedParaDate = CDate(objMasked.FormattedText)
    End If

End Function

Function StrParaMasked(objMasked As Object, sString As String)
'coloca sString em MaskedEdit verificando promptinclude

    If objMasked.PromptInclude = True Then
        
        objMasked.PromptInclude = False
        objMasked.Text = sString
        objMasked.PromptInclude = True
    
    Else
        
        objMasked.Text = sString
    
    End If
    
End Function

Sub Combo_Seleciona_ItemData(objComboBox As Object, ByVal lItemData As Long)
'seleciona o item da combo que tem ItemData igual ao passado por parametro

    Call ListOuCombo_Seleciona_ItemData(objComboBox, lItemData)
    
End Sub

Sub List_Seleciona_ItemData(objListBox As Object, ByVal lItemData As Long)
'seleciona o item da listbox que tem ItemData igual ao passado por parametro

    Call ListOuCombo_Seleciona_ItemData(objListBox, lItemData)
    
End Sub

Private Sub ListOuCombo_Seleciona_ItemData(objBox As Object, ByVal lItemData As Long)
'seleciona o item da combo ou listbox
Dim lErro As Long, iIndice As Integer
On Error GoTo Erro_ListOuCombo_Seleciona_ItemData

    'procura um item com o iTemData informado
    For iIndice = 0 To objBox.ListCount - 1

        If objBox.ItemData(iIndice) = lItemData Then
            objBox.ListIndex = iIndice
            Exit For
        End If
        
    Next

    Exit Sub

Erro_ListOuCombo_Seleciona_ItemData:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142348)

    End Select

    Exit Sub

End Sub

Private Function ListOuCombo_Item_Igual(objBox As Object, sString As String) As Long
'Seleciona ítem na listobox ou ComboBox referente a sString passada por parametro
'ATENÇÃO: Esta função ao selecionar ativa o ClickEvent

Dim lErro As Long
Dim iIndice As Integer
Dim iSucesso As Integer

On Error GoTo Error_ListOuCombo_Item_Igual

    iSucesso = 0

    'percorre todos os elementos da Box
    For iIndice = 0 To objBox.ListCount - 1

        'compara se codigo ja existe na Box
        If objBox.List(iIndice) = sString Then

            iSucesso = 1
            
            'Seleciona o ítem
            objBox.ListIndex = iIndice
            Exit For

        End If

    Next

    If iSucesso <> 1 Then Error 12253

    ListOuCombo_Item_Igual = SUCESSO

Exit Function

Error_ListOuCombo_Item_Igual:

    ListOuCombo_Item_Igual = Err

    Select Case Err

        Case 12253 'nao encontrou

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142349)

    End Select

    Exit Function

End Function

Private Function ListOuCombo_Item_Igual_Remove(objBox As Object, sString As String) As Long
'Remove ítem na ListBox ou ComboBox referente a sString passada por parametro

Dim lErro As Long
Dim iIndice As Integer
Dim iSucesso As Integer

On Error GoTo Error_ListOuCombo_Item_Igual_Remove

    iSucesso = 0

    'percorre todos os elementos da ComboBox
    For iIndice = 0 To objBox.ListCount - 1

        'compara se codigo ja existe na ComboBox
        If objBox.List(iIndice) = sString Then

            iSucesso = 1
            
            'remove o ítem da Combo
            objBox.RemoveItem (iIndice)
            Exit For

        End If

    Next

    If iSucesso <> 1 Then Error 19370

    ListOuCombo_Item_Igual_Remove = SUCESSO

Exit Function

Error_ListOuCombo_Item_Igual_Remove:

    ListOuCombo_Item_Igual_Remove = Err

    Select Case Err

        Case 19370 'nao encontrou

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142350)

    End Select

    Exit Function

End Function

Function SeguraFoco(objControle As Object, Frames As Object, iFrameAtual As Integer, Opcao As Object) As Long
'Segura o foco em objControle quando o Controle está em uma Frame de Frames dentro do TabStrip Opcao
'Frames deve se referir às Frames controladas pelo TabStrip
    
Dim lErro As Long
Dim objFrame As Frame  'Será a Frame (ligada a TabStrip) que contém o Controle
    
On Error GoTo Erro_SeguraFoco
    
    'Testa a profundidade do Frame em relação ao Controle
    If objControle.Container.Container <> "Frame" Then
        Set objFrame = objControle.Container
    ElseIf objControle.Container.Container.Container <> "Frame" Then
        Set objFrame = objControle.Container.Container
    Else
        Set objFrame = objControle.Container.Container.Container
    End If
    
    'Se FrameAtual não é o do controle
    If objFrame.Index <> iFrameAtual Then
       
       'Torna Frame do controle visível
        objFrame.Visible = True
        Frames(iFrameAtual).Visible = False
        iFrameAtual = objFrame.Index
        
        'Torna Tab da Frame do Controle o Tab selecionado
        Set Opcao.SelectedItem = Opcao.Tabs(objFrame.Index)
     
    End If
    
    'Segura o foco
    objControle.SetFocus
    
    SeguraFoco = SUCESSO
    
    Exit Function
    
Erro_SeguraFoco:

    SeguraFoco = Err

    lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142351)

    Exit Function

End Function


Function Data_Aumenta(sData As String) As Long

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_Data_Aumenta
    
    'Se está vazia deixa como está
    If sData <> "  /  /  " And sData <> "  /  /    " Then
    
        'Critica data
        If Not IsDate(sData) Then Error 6174
    
        'Aumenta data de um dia
        dtData = CDate(sData)
        dtData = DateAdd("d", 1, dtData)
        If Len(sData) = 10 Then
            sData = Format(dtData, "dd/mm/yyyy")
        Else
            sData = Format(dtData, "dd/mm/yy")
        End If
        
    End If
    
    Data_Aumenta = SUCESSO
    
    Exit Function
    
Erro_Data_Aumenta:

    Data_Aumenta = Err

    Select Case Err

        Case 6174 'Vazio porque o erro já foi tratado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142352)

    End Select

    Exit Function

End Function
Function Data_Diminui(sData As String) As Long

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_Data_Diminui
    
    'Se está vazia deixa como está
    If sData <> "  /  /  " And sData <> "  /  /    " Then
    
        'Critica data
        If Not IsDate(sData) Then Error 6175
    
        'Aumenta data de um dia
        dtData = CDate(sData)
        dtData = DateAdd("d", -1, dtData)
        If Len(sData) = 10 Then
            sData = Format(dtData, "dd/mm/yyyy")
        Else
            sData = Format(dtData, "dd/mm/yy")
        End If
        
    End If
    
    Data_Diminui = SUCESSO
    
    Exit Function
    
Erro_Data_Diminui:

    Data_Diminui = Err

    Select Case Err

        Case 6175 'Vazio porque o erro já foi tratado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142353)

    End Select

    Exit Function

End Function

Function Data_Up_Down_Click(MaskEdData As Object, ByVal iOpcao As Integer) As Long
'recebe o MaskEdBox de uma Data e a opção ( AUMENTA_DATA ou DIMINUI_DATA )

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_Data_Up_Down_Click

    MaskEdData.SetFocus
    
    'verifica se a data está vazia
    If Len(MaskEdData.ClipText) > 0 Then

        sData = MaskEdData.Text

        If iOpcao = AUMENTA_DATA Then

            lErro = Data_Aumenta(sData)
            If lErro <> SUCESSO Then Error 13064

        Else

            lErro = Data_Diminui(sData)
            If lErro <> SUCESSO Then Error 13041

        End If

        'exibe o novo valor da data
        MaskEdData.PromptInclude = False
        MaskEdData.Text = sData
        MaskEdData.PromptInclude = True

    End If

    Data_Up_Down_Click = SUCESSO

    Exit Function

Erro_Data_Up_Down_Click:

    Data_Up_Down_Click = Err

    Select Case Err

        Case 13041

        Case 13064

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142354)

    End Select

    Exit Function

End Function

Function Codigo_Extrai(sString As String) As Integer

Dim iPosicaoSeparador As Integer, sBuffer As String, dCodigo As Double
On Error GoTo Erro_Codigo_Extrai

    If Len(sString) = 0 Then
    
        Codigo_Extrai = 0
        
    Else
        iPosicaoSeparador = InStr(sString, SEPARADOR)
        If iPosicaoSeparador > 1 Then
            sBuffer = Mid(sString, 1, iPosicaoSeparador - 1)
        Else
            sBuffer = sString
        End If
        
        If IsNumeric(sBuffer) = False Then Error 7717
            
        Codigo_Extrai = CInt(sBuffer)
        
    End If
    
    Exit Function

Erro_Codigo_Extrai:

    Codigo_Extrai = 0

    Exit Function

End Function

Function LCodigo_Extrai(sString As String) As Long

Dim iPosicaoSeparador As Integer, sBuffer As String, dCodigo As Double
On Error GoTo Erro_LCodigo_Extrai

    iPosicaoSeparador = InStr(sString, SEPARADOR)
    If iPosicaoSeparador > 1 Then
        sBuffer = Mid(sString, 1, iPosicaoSeparador - 1)
    Else
        sBuffer = sString
    End If
    
    If IsNumeric(sBuffer) = False Then Error 7718
        
    LCodigo_Extrai = CLng(sBuffer)

    Exit Function

Erro_LCodigo_Extrai:

    LCodigo_Extrai = 0

    Exit Function

End Function

Function SCodigo_Extrai(sString As String) As String

Dim iPosicaoSeparador As Integer, sBuffer As String
On Error GoTo Erro_SCodigo_Extrai

    iPosicaoSeparador = InStr(sString, SEPARADOR)
    If iPosicaoSeparador > 1 Then
        sBuffer = Mid(sString, 1, iPosicaoSeparador - 1)
    Else
        sBuffer = sString
    End If
    
    SCodigo_Extrai = sBuffer

    Exit Function

Erro_SCodigo_Extrai:

    SCodigo_Extrai = ""

    Exit Function

End Function

Function Combo_Item_Igual_CI(objComboBox As Object)
'Seleciona ítem na ComboBox referente ao objComboBox.text passado por parametro
'Seleciona independente de ser digitada letra maiuscula ou letra minuscula
'ATENÇÃO: Esta função ao selecionar ativa o ClickEvent

    Combo_Item_Igual_CI = ListOuCombo_Item_Igual_CI(objComboBox, Trim(objComboBox.Text))

End Function

Private Function ListOuCombo_Item_Igual_CI(objBox As Object, sString As String) As Long
'Seleciona ítem na listobox ou ComboBox referente a sString passada por parametro
'Seleciona independente de ser digitada letra maiuscula ou letra minuscula
'ATENÇÃO: Esta função ao selecionar ativa o ClickEvent

Dim lErro As Long
Dim iIndice As Integer
Dim iSucesso As Integer

On Error GoTo Error_ListOuCombo_Item_Igual_CI

    iSucesso = 0

    'percorre todos os elementos da Box
    For iIndice = 0 To objBox.ListCount - 1

        'compara se codigo ja existe na Box
        If UCase(objBox.List(iIndice)) = UCase(sString) Then

            iSucesso = 1
            
            'Seleciona o ítem
            objBox.ListIndex = iIndice
            Exit For

        End If

    Next

    If iSucesso <> 1 Then Error 58583

    ListOuCombo_Item_Igual_CI = SUCESSO

    Exit Function

Error_ListOuCombo_Item_Igual_CI:

    ListOuCombo_Item_Igual_CI = Err

    Select Case Err

        Case 58583 'nao encontrou

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142355)

    End Select

    Exit Function

End Function

Function Combo_Item_Igual(objComboBox As Object) As Long
'Seleciona ítem na ComboBox referente ao objComboBox.text passado por parametro
'ATENÇÃO: Esta função ao selecionar ativa o ClickEvent

    Combo_Item_Igual = ListOuCombo_Item_Igual(objComboBox, Trim(objComboBox.Text))

End Function

Function Combo_Item_Igual_Remove(objComboBox As Object) As Long
'Remove ítem na ComboBox referente ao objComboBox.text passado por parametro
'ATENÇÃO: Esta função ao selecionar ativa o ClickEvent

    Combo_Item_Igual_Remove = ListOuCombo_Item_Igual_Remove(objComboBox, Trim(objComboBox.Text))

End Function

Function List_Item_Igual(objListBox As Object, sString As String) As Long
'Seleciona ítem na ListBox referente a sString passada por parametro
'ATENÇÃO: Esta função ao selecionar ativa o ClickEvent

    List_Item_Igual = ListOuCombo_Item_Igual(objListBox, sString)

End Function

Function List_Item_Igual_Remove(objListBox As Object, sString As String) As Long
'remove ítem na ListBox referente a sString passada por parametro

    List_Item_Igual_Remove = ListOuCombo_Item_Igual_Remove(objListBox, sString)

End Function

'Seleciona o ítem da Combo correspondente ao Text que pode ser numérico ou string
Function Combo_Seleciona(objComboBox As Object, iCodigo As Integer) As Long
        
Dim objCombo As ComboBox
Dim iNumerico As Integer
Dim iEncontrado As Integer
Dim lErro As Long
Dim dNumero As Double
Dim iIndice As Integer
Dim sString As String
Dim iPosicao As String
    
On Error GoTo Erro_Combo_Seleciona
        
    Set objCombo = objComboBox

    'Verifica se é o código numérico
    iNumerico = IsNumeric(objCombo.Text)
    
    If iNumerico Then
        
        dNumero = CDbl(objCombo.Text)
    
    Else  'Não é número puro
        'Determina existência e posição do caracter separador
        iPosicao = InStr(objCombo.Text, SEPARADOR)
        
        If iPosicao = 0 Then    'Não existe separador
            
            iNumerico = False
        
        Else                    'Existe separador
            
            iNumerico = IsNumeric(left(objCombo.Text, iPosicao - 1))
            
            If iNumerico Then
                dNumero = left(objCombo.Text, iPosicao - 1)
            End If
            
        End If
    
    End If
    
    iEncontrado = 0
    If iNumerico Then   'é NÚMERO
    
        'Verifica se é número INTEIRO
        If dNumero <> Int(dNumero) Or dNumero > MAXIMO_INTEIRO Then Error 6729
        'Verifica se é número NAO NEGATIVO
        If dNumero < 0 Then Error 6742
        
        'dNumero é inteiro positivo
        iCodigo = dNumero

        For iIndice = 0 To objCombo.ListCount - 1

           'Verifica se código existe na List da ComboBox
           If objCombo.ItemData(iIndice) = iCodigo Then

                'Seleciona o ítem da List
                objCombo.ListIndex = iIndice
                iEncontrado = 1
                Exit For
                
            End If
                
        Next
        
        If iEncontrado = 0 Then Error 6730
        
    Else   'é STRING
        
        sString = Trim(objCombo.Text)

        'Verifica se a string está na List da Combo
        For iIndice = 0 To objCombo.ListCount - 1
                
            If UCase(Mid(objCombo.List(iIndice), 1 + InStr(objCombo.List(iIndice), SEPARADOR))) = UCase(sString) Then
                             
                'Seleciona o ítem da List
                objCombo.ListIndex = iIndice
                iCodigo = objCombo.ItemData(iIndice)
                iEncontrado = 1
                Exit For
                
            End If
                
        Next
        
        If iEncontrado = 0 Then
        
            'Verifica se a string está na List da Combo
            For iIndice = 0 To objCombo.ListCount - 1
                    
                If Mid(objCombo.List(iIndice), 1 + InStr(objCombo.List(iIndice), SEPARADOR)) Like sString & "*" Then
                                 
                    'Seleciona o ítem da List
                    objCombo.ListIndex = iIndice
                    iCodigo = objCombo.ItemData(iIndice)
                    iEncontrado = 1
                    Exit For
                    
                End If
                    
            Next
            
        End If

        If iEncontrado = 0 Then Error 6731
    

    End If
    
    Combo_Seleciona = SUCESSO
    
    Exit Function

Erro_Combo_Seleciona:

    Combo_Seleciona = Err
    
    Select Case Err

        Case 6729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INTEIRO", Err, CStr(dNumero))
            objCombo.SetFocus
            
        Case 6742
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NEGATIVO", Err, CStr(dNumero))
            objCombo.SetFocus
            
        Case 6730   'Código não encontrado. Será tratado na tela chamadora
        'Deverá ir até o BD pesquisar o objeto pelo código
        'Se encontrar coloca na tela, se não dá aviso

        Case 6731   'String não encontrada. Será tratado na tela chamadora
        'Deverá dar mensagem de erro: objeto com nome/descrição <nome> não encontrado
        'Segura o foco
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142356)

    End Select

    Exit Function

End Function

'Seleciona o ítem da Combo correspondente ao Text que pode ser numérico-long ou string
Function LCombo_Seleciona(objComboBox As Object, lCodigo As Long) As Long

Dim objCombo As ComboBox
Dim iNumerico As Integer
Dim iEncontrado As Integer
Dim lErro As Long
Dim dNumero As Double
Dim iIndice As Integer
Dim sString As String
Dim iPosicao As String
  
On Error GoTo Erro_LCombo_Seleciona
        
    Set objCombo = objComboBox

    'Verifica se é o código numérico
    iNumerico = IsNumeric(objCombo.Text)
    
    If iNumerico Then
        
        dNumero = CDbl(objCombo.Text)
    
    Else  'Não é número puro
        'Determina existência e posição do caracter separador
        iPosicao = InStr(objCombo.Text, SEPARADOR)
        
        If iPosicao = 0 Then    'Não existe separador
            
            iNumerico = False
        
        Else                    'Existe separador
            
            iNumerico = IsNumeric(left(objCombo.Text, iPosicao - 1))
            
            If iNumerico Then
                dNumero = left(objCombo.Text, iPosicao - 1)
            End If
            
        End If
    
    End If
    
    iEncontrado = 0
    If iNumerico Then   'é NÚMERO
    
        'Verifica se é número LONG
        If dNumero <> Int(dNumero) Or dNumero > MAXIMO_LONG Then Error 25637
        
        'Verifica se é número NAO NEGATIVO
        If dNumero < 0 Then Error 25638
        
        'dNumero é inteiro positivo
        lCodigo = dNumero

        For iIndice = 0 To objCombo.ListCount - 1

           'Verifica se código existe na List da ComboBox
           If objCombo.ItemData(iIndice) = lCodigo Then

                'Seleciona o ítem da List
                objCombo.ListIndex = iIndice
                iEncontrado = 1
                Exit For
                
            End If
                
        Next
        
        If iEncontrado = 0 Then Error 25639
        
    Else   'é STRING
        
        sString = Trim(objCombo.Text)

        'Verifica se a string está na List da Combo
        For iIndice = 0 To objCombo.ListCount - 1
                
            If UCase(Mid(objCombo.List(iIndice), 1 + InStr(objCombo.List(iIndice), SEPARADOR))) = UCase(sString) Then
                             
                'Seleciona o ítem da List
                objCombo.ListIndex = iIndice
                lCodigo = objCombo.ItemData(iIndice)
                iEncontrado = 1
                Exit For
                
            End If
                
        Next
        
        If iEncontrado = 0 Then
        
            'Verifica se a string está na List da Combo
            For iIndice = 0 To objCombo.ListCount - 1
                    
                If Mid(objCombo.List(iIndice), 1 + InStr(objCombo.List(iIndice), SEPARADOR)) Like sString & "*" Then
                                 
                    'Seleciona o ítem da List
                    objCombo.ListIndex = iIndice
                    lCodigo = objCombo.ItemData(iIndice)
                    iEncontrado = 1
                    Exit For
                    
                End If
                    
            Next
            
        End If

        If iEncontrado = 0 Then Error 25640
    

    End If
    
    LCombo_Seleciona = SUCESSO
    
    Exit Function

Erro_LCombo_Seleciona:

    LCombo_Seleciona = Err
    
    Select Case Err

        Case 25637
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_LONG", Err, CStr(dNumero))
            objCombo.SetFocus
            
        Case 25638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NEGATIVO", Err, CStr(dNumero))
            objCombo.SetFocus
            
        Case 25639   'Código não encontrado. Será tratado na tela chamadora
        'Deverá ir até o BD pesquisar o objeto pelo código
        'Se encontrar coloca na tela, se não dá aviso

        Case 25640   'String não encontrada. Será tratado na tela chamadora
        'Deverá dar mensagem de erro: objeto com nome/descrição <nome> não encontrado
        'Segura o foco
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142357)

    End Select

    Exit Function

End Function

'Seleciona o ítem da ListBox correspondente a sTextBox (MaskedEdit ou TextBox) que pode ser numérico ou string
Function List_Seleciona(objMaskedEdit As Object, objListBox As Object, iCodigo As Integer) As Long
        
Dim objList As ListBox
Dim iNumerico As Integer
Dim iEncontrado As Integer
Dim lErro As Long
Dim dNumero As Double
Dim iIndice As Integer
Dim sString As String
Dim iPosicao As String
      
On Error GoTo Erro_List_Seleciona
        
    Set objList = objListBox
    
    'Testa se objeto é tipo MaskedEdit ou ListBox
    If Not (TypeName(objMaskedEdit) = "MaskEdBox" Or TypeName(objMaskedEdit) = "TextBox") Then Error 25632

    'Verifica se é o código numérico
    iNumerico = IsNumeric(objMaskedEdit.Text)
    
    If iNumerico Then
        
        dNumero = CDbl(objMaskedEdit.Text)
    
    Else  'Não é número puro
        'Determina existência e posição do caracter separador
        iPosicao = InStr(objMaskedEdit.Text, SEPARADOR)
        
        If iPosicao = 0 Then    'Não existe separador
            
            iNumerico = False
        
        Else                    'Existe separador
            
            iNumerico = IsNumeric(left(objMaskedEdit.Text, iPosicao - 1))
            
            If iNumerico Then
                dNumero = left(objMaskedEdit.Text, iPosicao - 1)
            End If
            
        End If
    
    End If
    
    iEncontrado = 0
    If iNumerico Then   'é NÚMERO
    
        'Verifica se é número INTEIRO
        If dNumero <> Int(dNumero) Or dNumero > MAXIMO_INTEIRO Then Error 25633
        'Verifica se é número NAO NEGATIVO
        If dNumero < 0 Then Error 25634
        
        'dNumero é inteiro positivo
        iCodigo = dNumero

        For iIndice = 0 To objList.ListCount - 1

           'Verifica se código existe na List da ComboBox
           If objList.ItemData(iIndice) = iCodigo Then

                'Seleciona o ítem da List
                objList.ListIndex = iIndice
                iEncontrado = 1
                Exit For
                
            End If
                
        Next
        
        If iEncontrado = 0 Then Error 25635
        
    Else   'é STRING
        
        sString = Trim(objMaskedEdit.Text)

        'Verifica se a string está na List da Combo
        For iIndice = 0 To objList.ListCount - 1
                
            If UCase(Mid(objList.List(iIndice), 1 + InStr(objList.List(iIndice), SEPARADOR))) = UCase(sString) Then
                             
                'Seleciona o ítem da List
                objList.ListIndex = iIndice
                iCodigo = objList.ItemData(iIndice)
                iEncontrado = 1
                Exit For
                
            End If
                
        Next
        
        If iEncontrado = 0 Then
        
            'Verifica se a string está na List da Combo
            For iIndice = 0 To objList.ListCount - 1
                    
                If Mid(objList.List(iIndice), 1 + InStr(objList.List(iIndice), SEPARADOR)) Like sString & "*" Then
                                 
                    'Seleciona o ítem da List
                    objList.ListIndex = iIndice
                    iCodigo = objList.ItemData(iIndice)
                    iEncontrado = 1
                    Exit For
                    
                End If
                    
            Next
            
        End If

        If iEncontrado = 0 Then Error 25636
    

    End If
    
    List_Seleciona = SUCESSO
    
    Exit Function

Erro_List_Seleciona:

    List_Seleciona = Err
    
    Select Case Err

        Case 25632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OBJETO_NAO_TEM_TIPO_ESPERADO", Err, objMaskedEdit.Name)
        
        Case 25633
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INTEIRO", Err)
            objMaskedEdit.SetFocus
            
        Case 25634
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NEGATIVO", Err)
            objMaskedEdit.SetFocus
            
        Case 25635   'Código não encontrado. Será tratado na tela chamadora
        'Deverá ir até o BD pesquisar o objeto pelo código
        'Se encontrar coloca na tela, se não dá aviso

        Case 25636   'String não encontrada. Será tratado na tela chamadora
        'Deverá dar mensagem de erro: objeto com nome/descrição <nome> não encontrado
        'Segura o foco
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142358)

    End Select

    Exit Function

End Function

'Seleciona o ítem da Combo correspondente ao Text que pode ser numérico ou string
'Para ser usada em Grids no lugar de Combo_Seleciona.
'Não segura foco em caso de erro.
Function Combo_Seleciona_Grid(objComboBox As Object, iCodigo As Integer) As Long
        
Dim objCombo As ComboBox
Dim iNumerico As Integer
Dim iEncontrado As Integer
Dim lErro As Long
Dim dNumero As Double
Dim iIndice As Integer
Dim sString As String
Dim iPosicao As String
    
On Error GoTo Erro_Combo_Seleciona_Grid
        
    Set objCombo = objComboBox

    'Verifica se é o código numérico
    iNumerico = IsNumeric(objCombo.Text)
    
    If iNumerico Then
        
        dNumero = CDbl(objCombo.Text)
    
    Else  'Não é número puro
        'Determina existência e posição do caracter separador
        iPosicao = InStr(objCombo.Text, SEPARADOR)
        
        If iPosicao = 0 Then    'Não existe separador
            
            iNumerico = False
        
        Else                    'Existe separador
            
            iNumerico = IsNumeric(left(objCombo.Text, iPosicao - 1))
            
            If iNumerico Then
                dNumero = left(objCombo.Text, iPosicao - 1)
            End If
            
        End If
    
    End If
    
    iEncontrado = 0
    If iNumerico Then   'é NÚMERO
    
        'Verifica se é número INTEIRO
        If dNumero <> Int(dNumero) Or dNumero > MAXIMO_INTEIRO Then Error 25083
        'Verifica se é número POSITIVO
        If dNumero < 1 Then Error 25084
        
        'dNumero é inteiro positivo
        iCodigo = dNumero

        For iIndice = 0 To objCombo.ListCount - 1

           'Verifica se código existe na List da ComboBox
           If objCombo.ItemData(iIndice) = iCodigo Then

                'Seleciona o ítem da List
                objCombo.ListIndex = iIndice
                iEncontrado = 1
                Exit For
                
            End If
                
        Next
        
        If iEncontrado = 0 Then Error 25085
        
    Else   'é STRING
        
        sString = Trim(objCombo.Text)

        'Verifica se a string está na List da Combo
        For iIndice = 0 To objCombo.ListCount - 1
                
            If UCase(objCombo.List(iIndice)) = UCase(sString) Or UCase(Mid(objCombo.List(iIndice), 1 + InStr(objCombo.List(iIndice), SEPARADOR))) = UCase(sString) Then
                             
                'Seleciona o ítem da List
                objCombo.ListIndex = iIndice
                iCodigo = objCombo.ItemData(iIndice)
                iEncontrado = 1
                Exit For
                
            End If
                
        Next
        
        If iEncontrado = 0 Then
        
            'Verifica se a string está na List da Combo
            For iIndice = 0 To objCombo.ListCount - 1
                    
                If objCombo.List(iIndice) = sString Or Mid(objCombo.List(iIndice), 1 + InStr(objCombo.List(iIndice), SEPARADOR)) Like sString & "*" Then
                                 
                    'Seleciona o ítem da List
                    objCombo.ListIndex = iIndice
                    iCodigo = objCombo.ItemData(iIndice)
                    iEncontrado = 1
                    Exit For
                    
                End If
                    
            Next
            
        End If

        If iEncontrado = 0 Then Error 25086
    

    End If
    
    Combo_Seleciona_Grid = SUCESSO
    
    Exit Function

Erro_Combo_Seleciona_Grid:

    Combo_Seleciona_Grid = Err
    
    Select Case Err

        Case 25083
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INTEIRO", Err)
                        
        Case 25084
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err)
                        
        Case 25085   'Código não encontrado. Será tratado na tela chamadora
        'Deverá ir até o BD pesquisar o objeto pelo código
        'Se encontrar coloca na tela, se não dá aviso

        Case 25086   'String não encontrada. Será tratado na tela chamadora
        'Deverá dar mensagem de erro: objeto com nome/descrição <nome> não encontrado
        'Segura o foco
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142359)

    End Select

    Exit Function

End Function

Function Combo_Item_Seleciona(objComboBox As Object) As Long
'Seleciona ítem na ComboBox referente ao codigo ( = 4 primeiros caracteres de objComboBox.text) passado por parametro

Dim lErro As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim iSucesso As Integer

On Error GoTo Error_Combo_Item_Seleciona

    iSucesso = 0

    iCodigo = Codigo_Extrai(Trim(objComboBox.Text))

    'percorre todos os elementos da ComboBox
    For iIndice = 0 To objComboBox.ListCount - 1

        'compara se codigo ja existe na ComboBox
        If objComboBox.ItemData(iIndice) = iCodigo Then

            iSucesso = 1

            'seleciona o item na ComboBox
            objComboBox.ListIndex = iIndice
            Exit For

        End If

    Next

    If iSucesso <> 1 Then Error 12250

    Combo_Item_Seleciona = SUCESSO

Exit Function

Error_Combo_Item_Seleciona:

    Combo_Item_Seleciona = Err

    Select Case Err

        Case 12250   'Não encontrou o ítem na List da Combo

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142360)

    End Select

    Exit Function

End Function

Function Preenche_Espacos(lCodigo As Long, sEspacos As String) As Long

Select Case lCodigo

    Case Is < 10
        sEspacos = "   "

    Case Is < 100
        sEspacos = "  "

    Case Is < 1000
        sEspacos = " "

    Case Else
        sEspacos = ""

End Select

End Function

Function Dias_Mes(iMes As Integer, iAno As Integer) As Integer
'Retorna o ultima dia do mes/ano passado como parametro

Dim iDiasMes As Integer

    'número de dias que o último mês possui
    Select Case iMes
        Case 2
            iDiasMes = 28
        Case 4, 6, 9, 11
            iDiasMes = 30
        Case Else
            iDiasMes = 31
    End Select

    If iMes = 2 Then
        If ((iAno Mod 4) = 0 And (iAno Mod 100 <> 0)) Or ((iAno Mod 400) = 0) Then iDiasMes = 29
    End If
        
    Dias_Mes = iDiasMes

End Function

Function TabStrip_PodeTrocarTab(iTabAtual As Integer, Opcao As Object, objForm As Object, Optional iTestaCTB As Integer = 1) As Long
'provoca o lost focus do controle ativo durante o evento click da TabStrip Opcao
' para fazer com que o tratamento de erro seja feita com o controle ainda visivel
' se isto gerar algum erro ou aviso a troca de frame é desfeita

Dim lErro As Long
Dim iFrameProximo As Integer

On Error GoTo Erro_TabStrip_PodeTrocarTab

''    If GL_lUltimoErro <> -9999 Then Error 20760
        
''    GL_lUltimoErro = 0
''
''    'para forçar lost focus
''    Opcao.SetFocus
''    DoEvents
''
''    'se ocorreu algum erro ou aviso
''    If GL_lUltimoErro <> 0 Then
''        Set Opcao.SelectedItem = Opcao.Tabs(iTabAtual)
''        Error 22884
''    End If
        
    If iTestaCTB = 1 Then
    
        If gcolModulo.Ativo(MODULO_CONTABILIDADE) <> MODULO_ATIVO Then
                
            If UCase(left(Opcao.SelectedItem.Caption, 6)) = UCase(TITULO_CONTABILIZACAO_RESUMIDO) Then
            
                Set Opcao.SelectedItem = Opcao.Tabs(iTabAtual)
                Error 19361
            
            End If
        
        End If
        
    End If
    
    If Not gcolModulo Is Nothing Then
    
        If gcolModulo.Ativo(MODULO_ESTOQUE) <> MODULO_ATIVO Then
                
            If UCase(left(Opcao.SelectedItem.Caption, 5)) = "ALMOX" Or UCase(left(Opcao.SelectedItem.Caption, 4)) = "DIST" Or UCase(Opcao.SelectedItem.Caption) = "RASTRO" Then
            
                Set Opcao.SelectedItem = Opcao.Tabs(iTabAtual)
                Error 20846
            
            End If
        
        End If
    
    End If
    
  ''''  GL_lUltimoErro = -9999
    
    TabStrip_PodeTrocarTab = SUCESSO

    Exit Function

Erro_TabStrip_PodeTrocarTab:

    TabStrip_PodeTrocarTab = Err

    Select Case Err

        Case 19361, 20760, 20846, 22884
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142361)

    End Select

''    GL_lUltimoErro = -9999
    
    Exit Function

End Function


''Function Produto_Formata(sProduto As String, sProdutoFormatado As String, iProdutoPreenchido As Integer) As Long
'''Devolve o produto no formato do BD: zeros a esquerda em segmentos numericos e brancos a direita para segmentos alfabeticos
'''A String de Entrada vem com os caracteres delimitadores e brancos(se for o caso). O comprimento dela é o comprimento total da máscara.
'''iProdutoPreenchido retornará PRODUTO_VAZIO se nenhum segmento do produto for preenchido,
'''caso contrário retornará PRODUTO_PREENCHIDO
''
''Dim lErro As Long
''Dim iNumSegmentos As Integer
''Dim iInd1 As Integer
''Dim iInd2 As Integer
''Dim objSegmento As New ClassSegmento
''Dim sProdutoBruto As String
''Dim sSegmentoBruto As String 'Segmento sendo formatado
''Dim sSegmentoFormatado As String 'Segmento sendo formatado
''Dim sCaracter As String 'caracter sendo tratado
''Dim iRestoVazio As Integer
''Dim iSegmentoVazio As Integer
''
''On Error GoTo Erro_Produto_Formata
''
''    iRestoVazio = SEGMENTOS_RESTANTES_CHEIOS
''    iProdutoPreenchido = PRODUTO_VAZIO
''    sProdutoBruto = sProduto
''    sProdutoFormatado = ""
''    iNumSegmentos = NumSegProduto()
''
''    For iInd1 = 1 To iNumSegmentos
''
''        lErro = SegmentoProduto(iInd1, objSegmento)
''        If lErro <> SUCESSO Then Error 6072
''
''        sSegmentoBruto = Left(sProdutoBruto, objSegmento.iTamanho)
''        sProdutoBruto = Mid(sProdutoBruto, objSegmento.iTamanho + Len(objSegmento.sDelimitador) + 1)
''        sSegmentoFormatado = ""
''
''        iSegmentoVazio = SEGMENTO_VAZIO
''
''        Select Case objSegmento.iTipo
''
''            Case SEGMENTO_NUMERICO
''
''                For iInd2 = 1 To objSegmento.iTamanho
''
''                    sCaracter = Mid(sSegmentoBruto, iInd2, 1)
''
''                    If sCaracter = " " Or sCaracter = "" Then
''                        sSegmentoFormatado = "0" & sSegmentoFormatado
''
''                    ElseIf sCaracter = "0" Then
''                        sSegmentoFormatado = sSegmentoFormatado & sCaracter
''
''                    ElseIf sCaracter = "-" Then
''                        Error 55692
''                    Else
''                        iSegmentoVazio = SEGMENTO_CHEIO
''                        sSegmentoFormatado = sSegmentoFormatado & sCaracter
''                    End If
''
''                Next
''
''            Case SEGMENTO_ALFANUMERICO, SEGMENTO_ASCII
''
''                For iInd2 = objSegmento.iTamanho To 1 Step -1
''
''                    sCaracter = Mid(sSegmentoBruto, iInd2, 1)
''
''                    If sCaracter = " " Or sCaracter = "" Then
''                        sSegmentoFormatado = sSegmentoFormatado & " "
''                    Else
''                        iSegmentoVazio = SEGMENTO_CHEIO
''                        sSegmentoFormatado = sCaracter & sSegmentoFormatado
''                    End If
''
''                Next
''
''        End Select
''
''        If iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS And iSegmentoVazio <> SEGMENTO_VAZIO Then Error 20736
''
''        If iSegmentoVazio = SEGMENTO_VAZIO Then
''            iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS
''        Else
''            iProdutoPreenchido = PRODUTO_PREENCHIDO
''        End If
''
''
''        sProdutoFormatado = sProdutoFormatado & sSegmentoFormatado
''
''    Next
''
''    If iProdutoPreenchido = PRODUTO_VAZIO Then sProdutoFormatado = ""
''
''    Produto_Formata = SUCESSO
''
''    Exit Function
''
''Erro_Produto_Formata:
''
''    Produto_Formata = Err
''
''
''    Select Case Err
''
''        Case 20736
''            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEG_MEIO_NAO_PREENCHIDOS", Err)
''
''        Case 25035 'Vazio porque o erro já foi tratado
''
''        Case 55692
''            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEG_NUM_CARACTER_INVALIDO", Err)
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142362)
''
''    End Select
''
''    Exit Function
''
''End Function

Function Limpa_Relatorio(Frm As Object) As Long
'limpa as telas de parametros de relatórios

Dim iIndice As Integer
Dim lErro As Long
    
On Error GoTo Erro_Limpa_Relatorio

    Call Limpa_Tela(Frm)
    
    With Frm
        For iIndice = 0 To .Controls.Count - 1
            If TypeName(.Controls(iIndice)) = "ComboBox" Then
                .Controls(iIndice).ListIndex = -1
            End If
        Next
    End With

    Exit Function
    
    Limpa_Relatorio = SUCESSO
    
Erro_Limpa_Relatorio:

    Limpa_Relatorio = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142363)

    End Select

    Exit Function
    
End Function

Function RelOpcoes_Testa_Combo(ComboBox As Object, sNome As String)
'verifica se sNOme está na combobox. Se não estiver insere. Usado na interface de relatórios

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_RelOpcoes_Testa_Combo:

    For iIndice = 0 To ComboBox.ListCount - 1
        If ComboBox.List(iIndice) = sNome Then Exit Function
    Next
        
    ComboBox.AddItem sNome
        
    RelOpcoes_Testa_Combo = SUCESSO
    
    Exit Function
    
Erro_RelOpcoes_Testa_Combo:

    RelOpcoes_Testa_Combo = Err
    
        Select Case Err
        
            Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142364)

    End Select

    Exit Function
    
End Function

Function RelOpcoes_ComboOpcoes_Preenche(objRelatorio As AdmRelatorio, ComboOpcoes As Object, objRelOpcoes As AdmRelOpcoes, objTela As Object) As Long

Dim colRelParametros As New Collection
Dim lErro As Long, sCodRel As String, iIndice As Integer
Dim objRelOpcoesAux As AdmRelOpcoes

On Error GoTo Erro_RelOpcoes_ComboOpcoes_Preenche

    sCodRel = objRelatorio.sCodRel
    
    'le os nomes das opcoes do relatório existentes no BD
    lErro = CF("RelOpcoes_Le_Todos", sCodRel, colRelParametros)
    If lErro <> SUCESSO Then Error 48592

    'preenche o ComboBox com os nomes das opções do relatório
    For Each objRelOpcoesAux In colRelParametros
        ComboOpcoes.AddItem objRelOpcoesAux.sNome
    Next

    'verifica se o nome da opção passada está no ComboBox
    For iIndice = 0 To ComboOpcoes.ListCount - 1

        If ComboOpcoes.List(iIndice) = objRelOpcoes.sNome Then

            ComboOpcoes.Text = ComboOpcoes.List(iIndice)
            Call objTela.PreencherParametrosNaTela(objRelOpcoes)

            Exit For

        End If

    Next
    
    RelOpcoes_ComboOpcoes_Preenche = SUCESSO

    Exit Function

Erro_RelOpcoes_ComboOpcoes_Preenche:

    RelOpcoes_ComboOpcoes_Preenche = Err

    Select Case Err

        Case 48592

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142365)

    End Select

    Exit Function

End Function

Sub RelOpcoes_ComboOpcoes_Click(objRelOpcoes As AdmRelOpcoes, ComboOpcoes As Object, objTela As Object)

Dim lErro As Long

On Error GoTo Erro_RelOpcoes_ComboOpcoes_Click

    If ComboOpcoes.ListIndex = -1 Then Exit Sub

    objRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Le", objRelOpcoes)
    If (lErro <> SUCESSO) Then Error 47726

    lErro = objTela.PreencherParametrosNaTela(objRelOpcoes)
    If lErro <> SUCESSO Then Error 47727

    Exit Sub

Erro_RelOpcoes_ComboOpcoes_Click:

    Select Case Err

        Case 47726, 47727

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142366)

    End Select

    Exit Sub

End Sub

Function PrimeiroDia_Semana(dtData As Date, dtDataPrimeiroDia As Date) As Long
'Esta funcao retorna o primeiro dia da semana da data em questao
'"A DATA DA SEGUNDA-FEIRA"

Dim iDiaSemana As Integer
Dim lErro As Long

On Error GoTo Erro_PrimeiroDia_Semana

    iDiaSemana = Weekday(dtData)

    Select Case iDiaSemana

        Case vbSunday
            dtDataPrimeiroDia = dtData + 1
        Case vbMonday
            dtDataPrimeiroDia = dtData
        Case vbTuesday
            dtDataPrimeiroDia = dtData - 1
        Case vbWednesday
            dtDataPrimeiroDia = dtData - 2
        Case vbThursday
            dtDataPrimeiroDia = dtData - 3
        Case vbFriday
            dtDataPrimeiroDia = dtData - 4
        Case vbSaturday
            dtDataPrimeiroDia = dtData - 5
        
        End Select

    PrimeiroDia_Semana = SUCESSO

    Exit Function

Erro_PrimeiroDia_Semana:

    PrimeiroDia_Semana = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142367)

        End Select

    Exit Function

End Function

Sub RelOpcoes_ComboOpcoes_Validate(ComboOpcoes As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_RelOpcoes_ComboOpcoes_Validate

    If ComboOpcoes.ListIndex <> -1 Then Exit Sub

    If Len(ComboOpcoes.Text) <> 0 Then Call Combo_Item_Igual(ComboOpcoes)
    
    Exit Sub

Erro_RelOpcoes_ComboOpcoes_Validate:

    Cancel = True
    
    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142368)

    End Select

    Exit Sub

End Sub

Sub MaskEdBox_TrataGotFocus(objMasked As Object, Optional viAlterado As Variant)
Dim sMascaraAntiga As String, iAlteradoAntigo As Integer

    If objMasked.Mask <> "" Then
    
        If InStr(objMasked.Mask, "#") <> 0 Or InStr(objMasked.Mask, "9") <> 0 Then
        
            If Len(objMasked.ClipText) = 0 Then
            
                If Not IsMissing(viAlterado) Then
                    iAlteradoAntigo = viAlterado
                End If
                
                sMascaraAntiga = objMasked.Mask
                objMasked.Mask = ""
                objMasked.Text = ""
                objMasked.Mask = sMascaraAntiga
                
                If Not IsMissing(viAlterado) Then
                    viAlterado = iAlteradoAntigo
                End If
            
            End If
            
        End If
        
    End If
    
End Sub


Function FormataCpoTexto(sTexto As String, iTam As Integer) As String
'formata campo texto completando com brancos à direita

Dim iTexto As Integer

    iTexto = Len(sTexto)

    If iTexto >= iTam Then

        FormataCpoTexto = Mid$(sTexto, 1, iTam)

    Else

        FormataCpoTexto = sTexto & Space(iTam - iTexto)

    End If


End Function

Function FormataCpoValor(vData As Variant, iTam As Integer) As String
'formata campo Valor alinhado-o à direita sem ponto e decimais, com zeros a esquerda

Dim sValor As String
Dim iTamanhoValor As Integer
Dim dValor As Double
Dim sDecimais As String
Dim sInteiro As String
Dim sNumero As String
Dim sEspacos As String
Dim iData As Integer

    sValor = Format(vData, "Fixed")
    sDecimais = Mid(sValor, Len(sValor) - 1, 2)
    sInteiro = Mid(sValor, 1, Len(sValor) - 3)
    sNumero = sInteiro & sDecimais

    iData = iTam - Len(sNumero)
    sEspacos = String(iData, "0")

    FormataCpoValor = sEspacos & sInteiro & sDecimais

End Function

Function FormataCpoNum(vData As Variant, iTam As Integer) As String
'formata campo numerico alinhado-o à direita sem ponto e decimais, com zeros a esquerda

Dim iData As Integer
Dim sData As String

    If Len(vData) = iTam Then

        FormataCpoNum = vData
        Exit Function

    End If

    iData = iTam - Len(vData)
    sData = String(iData, "0")

    FormataCpoNum = sData & vData

End Function

Function Formata_Cep(vCep As Variant, iTamanho As Integer) As String
'Formatao Cep passado completando com zeros no final até atingir o tamanho indicado

Dim iDiferenca As Integer

    iDiferenca = iTamanho - Len(vCep)
    
    If iDiferenca > 0 Then
        Formata_Cep = vCep & FormataCpoNum("", iDiferenca)
    Else
        Formata_Cep = vCep
    End If

    Exit Function

End Function

Sub TabStrip_TrataBeforeClick(Cancel As Integer, Opcao As Object)
''    GL_lUltimoErro = 0
''    SendKeys "{TAB}", True
''    If GL_lUltimoErro <> 0 Then Cancel = True
End Sub

Sub Formata_String_Numero(sStringRecebe As String, sStringRetorna As String)
    
Dim iTamanho As Integer
Dim sCaracter As String
Dim iIndice As Integer

    iTamanho = Len(Trim(sStringRecebe))
        
    For iIndice = 1 To iTamanho
                
        sCaracter = Mid(sStringRecebe, iIndice, 1)
        
        If IsNumeric(sCaracter) Then
            sStringRetorna = sStringRetorna & sCaracter
        End If
    
    Next
    
End Sub

Function FormataCpoData(dtData As Date) As String
'formata campo data como AAAAMMDD

    FormataCpoData = Format(dtData, "yyyymmdd")

End Function

Function FormataCpoQuantidade(vData As Variant, iTam As Integer) As String
'formata campo Quantidade alinhado-o à direita sem ponto e decimais, com zeros a esquerda

Dim sValor As String
Dim iTamanhoValor As Integer
Dim dValor As Double
Dim sDecimais As String
Dim sInteiro As String
Dim sNumero As String
Dim sEspacos As String
Dim iData As Integer

    sValor = Format(vData, "#,##0.000")
    sDecimais = Mid(sValor, Len(sValor) - 2, 3)
    sInteiro = Mid(sValor, 1, Len(sValor) - 4)
    sNumero = sInteiro & sDecimais

    iData = iTam - Len(sNumero)
    sEspacos = String(iData, "0")

    FormataCpoQuantidade = sEspacos & sInteiro & sDecimais

End Function

Function Calcula_Periodicidade(iCodigoPer As Integer, dtData As Date) As Date
'Retorna em dtData o utlimo dia do periodo que começa em dtData (quando é chamada)

Dim dtDataAux As Date
    
    Select Case iCodigoPer

        Case PERIODICIDADE_LIVRE
            dtDataAux = dtData
            
        Case PERIODICIDADE_SEMANAL
            dtDataAux = DateAdd("d", 7, dtData) - 1
            
        Case PERIODICIDADE_DECENDIAL
            If Day(dtData) >= 21 Then
                dtDataAux = DateAdd("d", 10, dtData)
            Else
                dtDataAux = DateAdd("d", 10, dtData) - 1
            End If
            
            'se ultrapassou o mes
            If Month(dtDataAux) <> Month(dtData) Then
                dtDataAux = Dias_Mes(Month(dtData), Year(dtData)) & Mid(dtData, 3)
            End If
            
        Case PERIODICIDADE_QUINZENAL
            If Day(dtData) >= 16 Then
                dtDataAux = DateAdd("d", 15, dtData)
            Else
                dtDataAux = DateAdd("d", 15, dtData) - 1
            End If
            
            'se ultrapassou o mes
            If Month(dtDataAux) <> Month(dtData) Then
                dtDataAux = Dias_Mes(Month(dtData), Year(dtData)) & Mid(dtData, 3)
            End If

        Case PERIODICIDADE_MENSAL
            dtDataAux = DateAdd("m", 1, dtData) - 1

        Case PERIODICIDADE_BIMESTRAL
            dtDataAux = DateAdd("m", 2, dtData) - 1

        Case PERIODICIDADE_TRIMESTRAL
            dtDataAux = DateAdd("m", 3, dtData) - 1

        Case PERIODICIDADE_QUADRIMESTRAL
            dtDataAux = DateAdd("m", 4, dtData) - 1

        Case PERIODICIDADE_SEMESTRAL
            dtDataAux = DateAdd("m", 6, dtData) - 1

        Case PERIODICIDADE_ANUAL
            dtDataAux = DateAdd("m", 12, dtData) - 1

    End Select

    Calcula_Periodicidade = dtDataAux

End Function

Public Function Formata_Campo(iAlinhamento As Integer, iTamanho As Integer, sCaracter As String, vCampo As Variant)
'Formata vCampo adicionando sCaracter até atingir iTamanho.
'sCaracter será concatenado ao lado direito ou esquerdo de acordo com iAlinhamento
'Retorna vCampo formatado em Formata_Campo
'Exemplo: iAlinhamento = ALINHAMENTO_ESQUERDA, iTamanho = 10, sCaracter = 0, vCampo= 10
'Formata_Campo = 0000000010

Dim sComplemento As String
Dim iCodASCII As Integer
Dim sCampo As String
Dim sComplemento1 As String

    'Se o campo não tem seu tamanho máximo
    If Len(vCampo) < iTamanho Then
    
        'Obtém o código ascii do caracter que será usado para complemetar o campo
        iCodASCII = Asc(sCaracter)
        
        'Monta o complemento do campo de acordo com o tamanho e o caracter passados como parâmetros
        sComplemento = String(iTamanho - Len(vCampo), iCodASCII)
        
    End If
    
    'Verifica se o complemento será posicionado a direita ou a esquerda do texto
    Select Case iAlinhamento
    
        'Se for para posicionar a esquerda
        Case ALINHAMENTO_ESQUERDA
            
            'Concatena o complemento ao campo, posicionando-o a esquerda e garantindo que o tamanho máximo não foi ultrapassado
            sCampo = left(sComplemento & vCampo, iTamanho)
        
        'Se for para posicionar a direita
        Case ALINHAMENTO_DIREITA
            
            'Concatena o complemento ao campo, posicionando-o a direita e garantindo que o tamanho máximo não foi ultrapassado
            sCampo = left(vCampo & sComplemento, iTamanho)
        
        'Se for para posicionar no centro
        Case ALINHAMENTO_CENTRALIZADO
        
            sComplemento1 = Mid(sComplemento, 1, (Len(sComplemento) \ 2))
            
            'Concatena o complemento ao campo, posicionando-o no centro e garantindo que o tamanho máximo não foi ultrapassado
            sCampo = CStr(Mid(sComplemento1 & vCampo, 1, Len(sComplemento1) + Len(vCampo))) & sComplemento1
            If Len(sCampo) < iTamanho Then sCampo = sCampo & left(sComplemento, 1)
                    
    End Select
    
    Formata_Campo = sCampo
    
    Exit Function
    
End Function

Public Function Formata_CNPJ_CPF(sTextoFormatado As String, ByVal sTextoOriginal As String) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_CNPJ_CPF

    Select Case Len(Trim(sTextoOriginal))

        Case STRING_CPF 'CPF

            sTextoFormatado = Format(sTextoOriginal, "000\.000\.000-00")

        Case STRING_CGC  'CGC

            sTextoFormatado = Format(sTextoOriginal, "00\.000\.000\/0000-00")

        Case Else

            sTextoFormatado = sTextoOriginal

    End Select

    Formata_CNPJ_CPF = SUCESSO

    Exit Function

Erro_Formata_CNPJ_CPF:

    Formata_CNPJ_CPF = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142369)

    End Select

    Exit Function

End Function

Public Function Arredonda_ParaCima(dQuantidadeOrig As Double) As Double

Dim dFixQuant As Double
Dim dQuantTemp As Double, dQuantidadeArr As Double

    dFixQuant = Fix(dQuantidadeOrig + 0.5 * Sgn(dQuantidadeOrig))
    dQuantTemp = dQuantidadeOrig - Int(dQuantidadeOrig)
    Select Case dQuantTemp
    
        Case 0
            dQuantidadeArr = dFixQuant
        Case 0.5
            dQuantidadeArr = dFixQuant
        Case Else
            If (dQuantTemp < 0.5) Then
                dFixQuant = dFixQuant + 1
                dQuantidadeArr = dFixQuant
            Else
                dQuantidadeArr = dFixQuant
            End If
    End Select

    Arredonda_ParaCima = dQuantidadeArr
    
End Function

Function DesacentuaTexto(ByVal sTexto As String) As String
'retorna uma copia do texto com a troca dos caracteres acentuados por nao acentuados

Dim iIndice As Integer, lErro As Long
Dim sCaracter As String
Dim sGuardaTexto As String
Dim iCodigo As Integer

On Error GoTo ErroDesacentuaTexto

    'Para cada Caracter do Texto
    For iIndice = 1 To Len(sTexto)
    
        'Seleciona caracter da posição iIndice
        sCaracter = Mid(sTexto, iIndice, 1)
        
        'Pega codigo ASC do caracter da selecionado acima
        iCodigo = Asc(sCaracter)
        
        'Verifica se caracter é acentuado
        Select Case iCodigo
        
            Case 192 To 197
                sCaracter = Chr(65)
            
            Case 199
                sCaracter = Chr(67)
            
            Case 200 To 203
                sCaracter = Chr(69)
            
            Case 204 To 207
                sCaracter = Chr(73)
            
            Case 210 To 214
                sCaracter = Chr(79)
            
            Case 217 To 220
                sCaracter = Chr(85)
            
            Case 224 To 229
                sCaracter = Chr(97)
            
            Case 231
                sCaracter = Chr(99)
            
            Case 232 To 235
                sCaracter = Chr(101)
            
            Case 236 To 239
                sCaracter = Chr(105)
            
            Case 242 To 246
                sCaracter = Chr(111)
            
            Case 249 To 252
                sCaracter = Chr(117)
        
        End Select
        
        sGuardaTexto = sGuardaTexto & sCaracter
    
    Next
    
    DesacentuaTexto = sGuardaTexto
    
    Exit Function
    
ErroDesacentuaTexto:
    
    DesacentuaTexto = sTexto
    
    Select Case gErr
    
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142370)
    
    End Select

    Exit Function

End Function

Function Combo_Obtem_ItemData(ByVal objBox As Object, ByVal sString As String, lItemData As Long) As Long
'Retorna o ItemData correspondente ao string

Dim lErro As Long
Dim iIndice As Integer
Dim iSucesso As Integer

On Error GoTo Erro_Combo_Obtem_ItemData

    iSucesso = 0

    'percorre todos os elementos da Box
    For iIndice = 0 To objBox.ListCount - 1

        'compara se codigo ja existe na Box
        If objBox.List(iIndice) = sString Then

            iSucesso = 1
        
            lItemData = objBox.ItemData(iIndice)
            
            Exit For

        End If

    Next

    If iSucesso <> 1 Then gError 12253

    Combo_Obtem_ItemData = SUCESSO
     
    Exit Function
    
Erro_Combo_Obtem_ItemData:

    Combo_Obtem_ItemData = gErr
     
    Select Case gErr
          
        Case 12253
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142371)
     
    End Select
     
    Exit Function

End Function

Function Combo_Obtem_String(ByVal objBox As Object, ByVal lItemData As Long, sString As String) As Long
'Retorna o string correspondente ao ItemData

Dim lErro As Long
Dim iIndice As Integer
Dim iSucesso As Integer

On Error GoTo Erro_Combo_Obtem_String

    iSucesso = 0

    'percorre todos os elementos da Box
    For iIndice = 0 To objBox.ListCount - 1

        'compara se codigo ja existe na Box
        If lItemData = objBox.ItemData(iIndice) Then

            iSucesso = 1
        
            sString = objBox.List(iIndice)
            
            Exit For

        End If

    Next

    If iSucesso <> 1 Then gError 12253

    Combo_Obtem_String = SUCESSO
     
    Exit Function
    
Erro_Combo_Obtem_String:

    Combo_Obtem_String = gErr
     
    Select Case gErr
          
        Case 12253
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142372)
     
    End Select
     
    Exit Function

End Function

Public Function Arredonda_Estoque(ByVal dQtde As Double) As Double

Dim iNumDigitos As Integer

    iNumDigitos = Len(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
    
    If dQtde >= 0 Then
        Arredonda_Estoque = Round(dQtde + 0.0000000001, iNumDigitos)
    Else
        Arredonda_Estoque = Round(dQtde - 0.0000000001, iNumDigitos)
    End If

End Function

'#########################################################
'Data em Hebraico
' This code demonstrates how to convert a Hebrew date into a
' Gregorian date. The code is written in VB.Net, but I purposely
' used very generic features so it would be easy to translate
' this into other languages. Also, I avoided using many
' optimizations in order to make the logic clearer.
'
' These functions assume that all the current rules of the
' Hebrew calendar were always in existence (which is not true)
' and all the current rules of the Gregorian calendar were always
' in existence (which is not true).

' Here is a very brief description of the Hebrew calendar.
'
' The Hebrew calendar is a lunisolar calendar.  This means that
' the months are in sync with the moon and the years stay in sync
' with the sun.  A solar year is approximately 365.25 days.  A
' lunar month is approximately 29.5 days.  Twelve lunar months is
' approximately 354 days (12 * 29.5=354).  Thus, a lunar year of
' twelve months is approximately 11.25 days shorter than the solar year.
' To make up for this shortfall, the Hebrew calendar adds a thirteenth
' month to seven years over a nineteen year period. Therefore, over
' a nineteen year period, the Hebrew calendar is approximately the
' same length as a nineteen year solar calendar.
'
' In order to understand this code, you must know the following
' terms:
'   Molad - new moon. Hebrew months start around the day of the
'           new moon
'   Chalakim - 1 / 1080 of an hour or 3 1/3 seconds
'   Tishrei - the first month of the Hebrew year (at least for
'             these calculations)
'   Rosh Hashanah - The Jewish new year which starts on Tishrei 1.
'
' The Hebrew calendar assumes the period of time between one new
' moon to the next is 29 days, 12 hours and 793 chalakim. The first
' molad after creation occurred on Monday, September, 7th -3760 at 5
' hours and 204 chalakim.  Technically, the Gregorian date would be
' in the year 3761 BCE because there was no year 0 in the Gregorian
' calendar, but we will use the year of -3760.

' Sample Usage:
'    ' Converts AdarB/7/5765 to 4/6/2005
'    MsgBox(HebCal.HebToGreg(5765, 7, 26))
'    ' Converts 4/6/2005 to AdarB/7/5765.  The function will return
'    ' a string of "07/26/5765" and nYearH, nMonthH, hDateH will get
'    ' set to 5765, 7, 26
'    MsgBox(HebCal.GregToHeb(#04/06/2005#, nYearH, nMonthH, nDateH))

'

' This function returns how many months there has been from the
' first Molad until the beginning of the year nYearH
Public Function MonSinceFirstMolad(ByVal nYearH As Integer) As Long
  
Dim nMonSinceFirstMolad As Long

    ' A shortcut to this function can simply be the following formula
    '   Return Int(((235 * nYearH) - 234) / 19)
    ' This formula is found in Remy Landau's website and he
    ' attributes it to Wolfgang Alexander Shochen. I will use a less
    ' optimized function which I believe shows the underlying logic
    ' better.

    ' count how many months there has been in all years up to last
    ' year. The months of this year hasn't happened yet.
    nYearH = nYearH - 1

    ' In the 19 year cycle, there will always be 235 months. That
    ' would be 19 years times 12 months plus 7 extra month for the
    ' leap years. (19 * 12) + 7 = 235.

    ' Get how many 19 year cycles there has been and multiply it by
    ' 235
    nMonSinceFirstMolad = Int(nYearH / 19) * 235
    ' Get the remaining years after the last complete 19 year cycle
    nYearH = nYearH Mod 19
    ' Add 12 months for each of those years
    nMonSinceFirstMolad = nMonSinceFirstMolad + (12 * nYearH)
    ' Add the extra months to account for the leap years
    If nYearH >= 17 Then
      nMonSinceFirstMolad = nMonSinceFirstMolad + 6
    ElseIf nYearH >= 14 Then
      nMonSinceFirstMolad = nMonSinceFirstMolad + 5
    ElseIf nYearH >= 11 Then
      nMonSinceFirstMolad = nMonSinceFirstMolad + 4
    ElseIf nYearH >= 8 Then
      nMonSinceFirstMolad = nMonSinceFirstMolad + 3
    ElseIf nYearH >= 6 Then
      nMonSinceFirstMolad = nMonSinceFirstMolad + 2
    ElseIf nYearH >= 3 Then
      nMonSinceFirstMolad = nMonSinceFirstMolad + 1
    End If
    
    MonSinceFirstMolad = nMonSinceFirstMolad
    
End Function

' This function returns if a given year is a leap year.
Public Function IsLeapYear(ByVal nYearH As Integer) As Boolean
    Dim nYearInCycle As Integer

    ' Find out which year we are within the cycle.  The 19th year of
    ' the cycle will return 0
    nYearInCycle = nYearH Mod 19
    
    IsLeapYear = nYearInCycle = 3 Or _
           nYearInCycle = 6 Or _
           nYearInCycle = 8 Or _
           nYearInCycle = 11 Or _
           nYearInCycle = 14 Or _
           nYearInCycle = 17 Or _
           nYearInCycle = 0
End Function

' This function figures out the Gregorian Date that corresponds to
' the first day of Tishrei, the first month of the Hebrew
' calendar, for a given Hebrew year.
Public Function Tishrei1(ByVal nYearH As Integer) As Date
    Dim nMonthsSinceFirstMolad As Long
    Dim nChalakim As Long
    Dim nHours As Long
    Dim nDays As Long
    Dim nDayOfWeek As Integer
    Dim dTishrei1 As Date

    ' We want to calculate how many days, hours and chalakim it has
    ' been from the time of 0 days, 0 hours and 0 chalakim to the
    ' molad at the beginning of year nYearH.
    '
    ' The period between one new moon to the next is 29 days, 12
    ' hours and 793 chalakim. We must multiply that by the amount
    ' of months that transpired since the first molad. Then we add
    ' the time of the first molad (Monday, 5 hours and 204 chalakim)
    nMonthsSinceFirstMolad = MonSinceFirstMolad(nYearH)
    nChalakim = 793 * nMonthsSinceFirstMolad
    nChalakim = nChalakim + 204
    ' carry the excess Chalakim over to the hours
    nHours = Int(nChalakim / 1080)
    nChalakim = nChalakim Mod 1080

    nHours = nHours + (nMonthsSinceFirstMolad * 12)
    nHours = nHours + 5
    ' carry the excess hours over to the days
    nDays = Int(nHours / 24)
    nHours = nHours Mod 24

    nDays = nDays + (29 * nMonthsSinceFirstMolad)
    nDays = nDays + 2

    ' figure out which day of the week the molad occurs.
    ' Sunday = 1, Moday = 2 ..., Shabbos = 0
    nDayOfWeek = nDays Mod 7

    ' In a perfect world, Rosh Hashanah would be on the day of the
    ' molad. The Hebrew calendar makes four exceptions where we
    ' push off Rosh Hashanah one or two days. This is done to
    ' prevent three situation. Without explaining why, the three
    ' situations are:
    '   1) We don't want Rosh Hashanah to come out on Sunday,
    '      Wednesday or Friday
    '   2) We don't want Rosh Hashanah to be on the day of the
    '      molad if the molad occurs after the beginning of 18th
    '      hour.
    '   3) We want to limit years to specific lengths.  For non-leap
    '      years, we limit it to either 353, 354 or 355 days.  For
    '      leap years, we limit it to either 383, 384 or 385 days.
    '      If setting Rosh Hashanah to the day of the molad will
    '      cause this year, or the previous year to fall outside
    '      these lengths, we push off Rosh Hashanah to get the year
    '      back to a valid length.
    ' This code handles these exceptions.

    If Not IsLeapYear(nYearH) And _
       nDayOfWeek = 3 And _
       (nHours * 1080) + nChalakim >= _
       (9 * 1080) + 204 Then
      ' This prevents the year from being 356 days. We have to push
      ' Rosh Hashanah off two days because if we pushed it off only
      ' one day, Rosh Hashanah would comes out on a Wednesday. Check
      ' the Hebrew year 5745 for an example.
      nDayOfWeek = 5
      nDays = nDays + 2
    ElseIf IsLeapYear(nYearH - 1) And _
           nDayOfWeek = 2 And _
           (nHours * 1080) + nChalakim >= _
           (15 * 1080) + 589 Then
      ' This prevents the previous year from being 382 days. Check
      ' the Hebrew Year 5766 for an example. If Rosh Hashanah was not
      ' pushed off a day then 5765 would be 382 days
      nDayOfWeek = 3
      nDays = nDays + 1
    Else
      ' see rule 2 above. Check the Hebrew year 5765 for an example
      If nHours >= 18 Then
        nDayOfWeek = nDayOfWeek + 1
        nDayOfWeek = nDayOfWeek Mod 7
        nDays = nDays + 1
      End If
      ' see rule 1 above. Check the Hebrew year 5765 for an example
      If nDayOfWeek = 1 Or _
         nDayOfWeek = 4 Or _
         nDayOfWeek = 6 Then
        nDayOfWeek = nDayOfWeek + 1
        nDayOfWeek = nDayOfWeek Mod 7
        nDays = nDays + 1
      End If
    End If

    ' Here we want to add nDays to creation
    '    dTishrie1 = creation + nDays
    ' Unfortunately, VB.Net doesn't handle negative years very well.
    ' I therefore picked a Random date (1/1/1900) and figured out how
    ' many days it is after the creation (2067025). Then I subtracted
    ' 2067025 from nDays.
    nDays = nDays - 2067025
    dTishrei1 = #1/1/1900# ' 2067025 days after creation
    
    dTishrei1 = DateAdd("d", nDays, dTishrei1)
    
    Tishrei1 = dTishrei1
    
  End Function

' This function gets the length of a Hebrew year.
Public Function LengthOfYear(ByVal nYearH As Integer) As Integer
    Dim dThisTishrei1 As Date
    Dim dNextTishrei1 As Date
    Dim diff As Integer

    ' subtract the date of this year from the date of next year
    dThisTishrei1 = Tishrei1(nYearH)
    dNextTishrei1 = Tishrei1(nYearH + 1)
    
    diff = -DateDiff("d", dNextTishrei1, dThisTishrei1)
    
    LengthOfYear = diff
    
  End Function

' This function converts a Hebrew date into the Gregorian date
' nYearH - is the Hebrew year
' nMonth - Tishrei=1
'          Cheshvan=2
'          Kislev=3
'          Teves=4
'          Shevat=5
'          Adar A=6 (only valid on leap years)
'          Adar=7   (Adar B for leap years)
'          Nisan=8
'          Iyar=9
'          Sivan=10
'          Tamuz=11
'          Av=12
'          Elul=13
Public Function HebToGreg(ByVal nYearH As Integer, _
                          ByVal nMonthH As Integer, _
                          ByVal nDateH As Integer) As Date
    Dim nLengthOfYear As Integer
    Dim bLeap As Boolean
    Dim dGreg As Date
    Dim nMonth As Integer
    Dim nMonthLen As Integer
    Dim bHaser As Boolean
    Dim bShalem As Boolean

    bLeap = IsLeapYear(nYearH)
    nLengthOfYear = LengthOfYear(nYearH)

    ' The regular length of a non-leap year is 354 days.
    ' The regular length of a leap year is 384 days.
    ' On regular years, the length of the months are as follows
    '   Tishrei (1)   30
    '   Cheshvan(2)   29
    '   Kislev  (3)   30
    '   Teves   (4)   29
    '   Shevat  (5)   30
    '   Adar A  (6)   30     (only valid on leap years)
    '   Adar    (7)   29     (Adar B for leap years)
    '   Nisan   (8)   30
    '   Iyar    (9)   29
    '   Sivan   (10)  30
    '   Tamuz   (11)  29
    '   Av      (12)  30
    '   Elul    (13)  29
    ' If the year is shorter by one less day, it is called a haser
    ' year. Kislev on a haser year has 29 days. If the year is longer
    ' by one day, it is called a shalem year. Cheshvan on a shalem
    ' year is 30 days.
    bHaser = nLengthOfYear = 353 Or nLengthOfYear = 383
    bShalem = nLengthOfYear = 355 Or nLengthOfYear = 385

    ' get the date for Tishrei 1
    dGreg = Tishrei1(nYearH)
    ' Now count up days within the year
    For nMonth = 1 To nMonthH - 1
      Select Case nMonth
        Case 1, 5, 8, 10, 12 ' 30 day months
          nMonthLen = 30
        Case 4, 7, 9, 11, 13 ' 29 day months
          nMonthLen = 29
        Case 6 ' There is only an Adar A on a leap years
          nMonthLen = IIf(bLeap, 30, 0)
        Case 2 ' Cheshvan, see note above
          nMonthLen = IIf(bShalem, 30, 29)
        Case 3 ' Kislev, see note above
          nMonthLen = IIf(bHaser, 29, 30)
      End Select
      dGreg = DateAdd("d", nMonthLen, dGreg)
    Next
    dGreg = DateAdd("d", nDateH - 1, dGreg)
    
    HebToGreg = dGreg
    
End Function

' This function converts a Gregorian date into the Hebrew date.  The
' function returns the hebrew month as a string in the format MM/DD/YYYY.
' Also, the parameters nYearH, nMonthH and hDateH, which are sent by
' reference, will get set the Hebrew year, month and date. See function
' HebToGreg() for the definition of the month numbers.
Public Function GregToHeb(ByVal dGreg As Date, _
                          ByRef nYearH As Integer, _
                          ByRef nMonthH As Integer, _
                          ByRef nDateH As Integer) As String
    Dim nOneMolad As Double
    Dim nAvrgYear As Double
    Dim nDays As Long
    Dim dTishrei1 As Date
    Dim nLengthOfYear As Integer
    Dim bLeap As Boolean
    Dim bHaser As Boolean
    Dim bShalem As Boolean
    Dim nMonthLen As Integer
    Dim bWhile As Boolean, sMesExtenso As String

    ' The basic algorythm to get Hebrew date for the Gregorian date dGreg.
    ' 1) Find out how many days dGreg is after creation.
    ' 2) Based on those days, estimate the Hebrew year
    ' 3) Now that we a good estimate of the Hebrew year, use brute force to
    '    find the Gregorian date for Tishrei 1 prior to or equal to dGreg
    ' 4) Add to Tishrei 1 the amount of days dGreg is after Tishrei 1

    ' Figure out how many days are in a month.
    ' 29 days + 12 hours + 793 chalakim
    nOneMolad = 29 + (12 / 24) + (793 / (1080 * 24))
    ' Figure out the average length of a year. The hebrew year has exactly
    ' 235 months over 19 years.
    nAvrgYear = nOneMolad * (235 / 19)
    ' Get how many days dGreg is after creation. See note as to why I
    ' use 1/1/1900 and add 2067025
    nDays = -DateDiff("d", dGreg, #1/1/1900#)
    nDays = nDays + 2067025 ' 2067025 days after creation

    ' Guess the Hebrew year. This should be a pretty accurate guess.
    nYearH = Int(CDbl(nDays) / nAvrgYear) + 1
    ' Use brute force to find the exact year nYearH. It is the Tishrei 1 in
    ' the year <= dGreg.
    dTishrei1 = Tishrei1(nYearH)
    If dTishrei1 = dGreg Then
      ' If we got lucky and landed on the exact date, we can stop here
      nMonthH = 1
      nDateH = 1
    Else
      ' Here is the brute force.  Either count up or count down nYearH
      ' until Tishrei 1 is <= dGreg.
      If dTishrei1 < dGreg Then
        ' If Tishrei 1, nYearH is less than dGreg, count nYearH up.
        Do While Tishrei1(nYearH + 1) <= dGreg
          nYearH = nYearH + 1
        Loop
      Else
        ' If Tishrei 1, nYearH is greater than dGreg, count nYearH down.
        nYearH = nYearH - 1
        Do While Tishrei1(nYearH) > dGreg
          nYearH = nYearH - 1
        Loop
      End If

      ' Subtract Tishrei 1, nYearH from dGreg. That should leave us with
      ' how many days we have to add to Tishrei 1
      nDays = -DateDiff("d", dGreg, Tishrei1(nYearH))

      ' Find out what type of year it is so that we know the length of the
      ' months
      nLengthOfYear = LengthOfYear(nYearH)
      bHaser = nLengthOfYear = 353 Or nLengthOfYear = 383
      bShalem = nLengthOfYear = 355 Or nLengthOfYear = 385
      bLeap = IsLeapYear(nYearH)

      ' Add nDays to Tishrei 1.
      nMonthH = 1
      Do
        Select Case nMonthH
          Case 1, 5, 8, 10, 12 ' 30 day months
            nMonthLen = 30
          Case 4, 7, 9, 11, 13 ' 29 day months
            nMonthLen = 29
          Case 6 ' Adar A (6) will be skipped on non-leap years
            nMonthLen = 30
          Case 2 ' Cheshvan, see note above
            nMonthLen = IIf(bShalem, 30, 29)
          Case 3 ' Kislev, see note above
            nMonthLen = IIf(bHaser, 29, 30)
        End Select
        If nDays >= nMonthLen Then
          bWhile = True
          If bLeap Or nMonthH <> 5 Then
            nMonthH = nMonthH + 1
          Else
            ' We can skip Adar A (6) if its not a leap year
            nMonthH = nMonthH + 2
          End If
          nDays = nDays - nMonthLen
        Else
          bWhile = False
        End If
      Loop While bWhile
      ' Add the remaining days to Date
      nDateH = nDays + 1
    End If
    
    Select Case nMonthH
        Case 1
            sMesExtenso = "Tishri"
        Case 2
            sMesExtenso = "Cheshvan"
        Case 3
            sMesExtenso = "Kislev"
        Case 4
            sMesExtenso = "Teveth"
        Case 5
            sMesExtenso = "Shevat"
        Case 6
            sMesExtenso = "Adar I"
        Case 7
            If bLeap Then
                sMesExtenso = "Adar II"
            Else
                sMesExtenso = "Adar"
            End If
        Case 8
            sMesExtenso = "Nisan"
        Case 9
            sMesExtenso = "Iyar"
        Case 10
            sMesExtenso = "Sivan"
        Case 11
            sMesExtenso = "Tamuz"
        Case 12
            sMesExtenso = "Av"
        Case 13
            sMesExtenso = "Elul"
        
    End Select
    
    GregToHeb = Format(nDateH, "00") & " " & sMesExtenso & " " & CStr(nYearH)
    
End Function
'#######################################################################

Function StrParaDblErr(sTexto As String) As Double
'retorna sTexto como double

On Error GoTo Erro_StrParaDblErr

    If Len(Trim(sTexto)) = 0 Then
        StrParaDblErr = 0
    Else
        StrParaDblErr = CDbl(sTexto)
    End If
    
    Exit Function

Erro_StrParaDblErr:

    Select Case gErr
    
        Case 6, 13
            StrParaDblErr = 0
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178423)
     
    End Select
     
    Exit Function
    
End Function

Function StrParaIntErr(sTexto As String) As Integer
'retorna sTexto como Integer

On Error GoTo Erro_StrParaIntErr

    If Len(Trim(sTexto)) = 0 Then
        StrParaIntErr = 0
    Else
        StrParaIntErr = CInt(sTexto)
    End If

    Exit Function

Erro_StrParaIntErr:

    Select Case gErr
    
        Case 6, 13
            StrParaIntErr = 0
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178424)
     
    End Select
     
    Exit Function

End Function

Function StrParaLongErr(sTexto As String) As Long
'retorna sTexto como Integer

On Error GoTo Erro_StrParaLongErr

    If Len(Trim(sTexto)) = 0 Then
        StrParaLongErr = 0
    Else
        StrParaLongErr = CLng(sTexto)
    End If

    Exit Function

Erro_StrParaLongErr:

    Select Case gErr
    
        Case 6, 13
            StrParaLongErr = 0
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178425)
     
    End Select
     
    Exit Function

End Function

Function StrParaDateErr(sTexto As String) As Date
'retorna texto como data

On Error GoTo Erro_StrParaDateErr

    If Len(sTexto) = 0 Or left(sTexto, 2) = "  " Then
        StrParaDateErr = DATA_NULA
    Else
        StrParaDateErr = CDate(sTexto)
        If StrParaDateErr < DATA_NULA Then StrParaDateErr = DATA_NULA '20/06/200' dá erro na execucao do sql
    End If
    
    Exit Function
    
Erro_StrParaDateErr:

    Select Case gErr
    
        Case 6, 13
            StrParaDateErr = 0
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178426)
     
    End Select
     
    Exit Function
    
End Function

Function StrParaMasked2(objMasked As Object, ByVal sString As String)
'coloca sString em MaskedEdit verificando promptinclude

Dim sString2 As String

    If objMasked.PromptInclude = True Then
        
        objMasked.PromptInclude = False
        If InStr(objMasked.Mask, "C") <> 0 Then
            sString2 = Replace(sString, ".", "")
        Else
            sString2 = sString
        End If
        objMasked.Text = sString2
        objMasked.PromptInclude = True
    
    Else
        
        objMasked.Text = sString
    
    End If
    
End Function


