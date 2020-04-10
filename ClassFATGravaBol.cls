VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Endereco_Grava_Customizado(ByVal objEndereco As ClassEndereco) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_Endereco_Grava_Customizado

    If Not (objEndereco.objInfoUsu Is Nothing) Then

        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 208206
        Next
    
        'Le a tabelaTRVOcorrencias
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Enderecos WHERE Codigo= ? ", 0, lCodigo, objEndereco.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 208207
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208208
    
        If lErro <> AD_SQL_SUCESSO Then gError 208209
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 208210
    
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Enderecos SET Skype= ?, Radio = ? ", alComando(0), objEndereco.objInfoUsu.sSkype, objEndereco.objInfoUsu.sRadio)
        If lErro <> AD_SQL_SUCESSO Then gError 208211
            
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If

    Endereco_Grava_Customizado = SUCESSO

    Exit Function

Erro_Endereco_Grava_Customizado:

    Endereco_Grava_Customizado = gErr

    Select Case gErr
        
        Case 208206
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208207, 208208
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ENDERECOS", gErr)

        Case 208209
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECOS_NAO_CADASTRADO", gErr)

        Case 208210
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ENDERECOS", gErr)

        Case 208211
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ENDERECOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208212)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
