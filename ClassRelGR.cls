VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gobjCompServ As ClassCompServ
Dim bAchou As Boolean

Public Function ObterCodOperacao(ByVal lCodigo As Long, sPorto As String) As Long
'Para obter o Código de Operação e emitir as perguntas necessárias

Dim lErro As Long

On Error GoTo Erro_ObterCodOperacao

    If gobjCompServ.lCodigo <> lCodigo Then
        Call ObterLeitura(lCodigo)
    End If

    sPorto = gobjCompServ.sPorto

    ObterCodOperacao = SUCESSO

    Exit Function

Erro_ObterCodOperacao:

    ObterCodOperacao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function ObterCodOperacao1(ByVal lCodigo As Long, sAlfandega As String) As Long
'Para obter o Código de Operação e emitir as perguntas necessárias

Dim lErro As Long

On Error GoTo Erro_ObterCodOperacao1

    If gobjCompServ.lCodigo <> lCodigo Then
        Call ObterLeitura(lCodigo)
    End If

    sAlfandega = gobjCompServ.sAlfandega

    ObterCodOperacao1 = SUCESSO

    Exit Function

Erro_ObterCodOperacao1:

    ObterCodOperacao1 = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function ObterCodOperacao2(ByVal lCodigo As Long, sMotorista As String) As Long
'Para obter o Código de Operação e emitir as perguntas necessárias

Dim lErro As Long
Dim objCompServItem As New ClassCompServItem

On Error GoTo Erro_ObterCodOperacao2

    If gobjCompServ.lCodigo <> lCodigo Then
        Call ObterLeitura(lCodigo)
    End If

    For Each objCompServItem In gobjCompServ.colCompServItem
        'e o nome reduzido do documento interno igual a "CTRC"
        If objCompServItem.sDocIntTipo = "CTRC" Then
            sMotorista = objCompServItem.sMotorista
            Exit For
        End If
    Next

    ObterCodOperacao2 = SUCESSO

    Exit Function

Erro_ObterCodOperacao2:

    ObterCodOperacao2 = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function ObterCodOperacao3(ByVal lCodigo As Long, sPlacaCaminhao As String) As Long
'Para obter o Código de Operação e emitir as perguntas necessárias

Dim lErro As Long
Dim objCompServItem As New ClassCompServItem

On Error GoTo Erro_ObterCodOperacao3

    If gobjCompServ.lCodigo <> lCodigo Then
        Call ObterLeitura(lCodigo)
    End If

    bAchou = False

    'Se não for mercado interno
    If gobjCompServ.iTipoOperacao <> 2 Then

        For Each objCompServItem In gobjCompServ.colCompServItem
            'e o nome reduzido do documento interno igual a "CTRC"
            If objCompServItem.sDocIntTipo = "CTRC" Then
                If Len(Trim(objCompServItem.sPlacaCaminhao)) > 0 Then
                    sPlacaCaminhao = objCompServItem.sPlacaCaminhao
                    bAchou = True
                End If
                Exit For
            End If
        Next

    End If

    ObterCodOperacao3 = SUCESSO

    Exit Function

Erro_ObterCodOperacao3:

    ObterCodOperacao3 = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function ObterCodOperacao4(ByVal lCodigo As Long, sPlacaCarreta As String) As Long
'Para obter o Código de Operação e emitir as perguntas necessárias

Dim lErro As Long
Dim objCompServItem As New ClassCompServItem

On Error GoTo Erro_ObterCodOperacao4

    If gobjCompServ.lCodigo <> lCodigo Then
        Call ObterLeitura(lCodigo)
    End If

    If bAchou = False Then

        'Se não for mercado interno
        If gobjCompServ.iTipoOperacao <> 2 Then

            For Each objCompServItem In gobjCompServ.colCompServItem
                'e o nome reduzido do documento interno igual a "CTRC"
                If objCompServItem.sDocIntTipo = "CTRC" Then
                    If Len(Trim(objCompServItem.sPlacaCarreta)) > 0 Then sPlacaCarreta = objCompServItem.sPlacaCarreta
                    Exit For
                End If
            Next

        End If

    End If

    ObterCodOperacao4 = SUCESSO

    Exit Function

Erro_ObterCodOperacao4:

    ObterCodOperacao4 = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Sub ObterLeitura(lCodigo As Long)
'Para obter o Código de Operação e emitir as perguntas necessárias

Dim lErro As Long
Dim objCompServItem As New ClassCompServItem

On Error GoTo Erro_ObterLeitura

    Set gobjCompServ = New ClassCompServ

    'Passa a chave para gobjCompServ
    gobjCompServ.iFilialEmpresa = giFilialEmpresa
    gobjCompServ.lCodigo = lCodigo

    'Le os dados da tabela Comprovante de Servico
    lErro = CF("CompServGR_Le", gobjCompServ)
    If lErro <> SUCESSO And lErro <> 97419 Then gError 99308

    'Se não encontrar --> Erro
    If lErro = 97419 Then gError 99309

    'Le os dados dos itens de serviço
    'relacionados com o servico que se relaciona com
    'o Comp Serv em questão
    lErro = CF("CompServGR_Le_CompServItem", gobjCompServ)
    If lErro <> SUCESSO Then gError 99310

    'Se for importação
    If gobjCompServ.iTipoOperacao = 0 Then

        For Each objCompServItem In gobjCompServ.colCompServItem
            'e o nome reduzido do doc. externo igual a "DI"
            If objCompServItem.sDocExtTipo = "DI" Then
                lErro = Chama_Tela_Modal("RelOpPorto", gobjCompServ)
                Exit For
            End If
            If objCompServItem.sDocExtTipo = "DTA" Then
              lErro = Chama_Tela_Modal("RelOpPortAlf", gobjCompServ)
              Exit For
            End If
        Next
    'Se for exportação, pega direto o Porto
    ElseIf gobjCompServ.iTipoOperacao = 1 Or gobjCompServ.iTipoOperacao = 3 Then
        lErro = Chama_Tela_Modal("RelOpPorto", gobjCompServ)
    End If

    Exit Sub

Erro_ObterLeitura:

    Select Case gErr

        Case 99308, 99310

        Case 99309
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_CADASTRADO", gErr, gobjCompServ.lCodigo, gobjCompServ.iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub Class_Initialize()
    Set gobjCompServ = New ClassCompServ
End Sub

Private Sub Class_Terminate()

    Set gobjCompServ = Nothing

End Sub
