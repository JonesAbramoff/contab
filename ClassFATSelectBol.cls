VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Endereco_Le_Customizado(ByVal objEndereco As ClassEndereco) As Long

Dim lErro As Long
Dim lComando As Long
Dim sSkype As String, sRadio As String
Dim objEnderecoBol As New ClassEnderecoBol

On Error GoTo Erro_Endereco_Le_Customizado

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208213

    'Alocação de espaço no buffer
    sSkype = String(STRING_SKYPE, 0)
    sRadio = String(STRING_RADIO, 0)

    'Le a tabelaTRVOcorrencias
    lErro = Comando_Executar(lComando, "SELECT Skype, Radio FROM Enderecos WHERE Codigo= ? ", sSkype, sRadio, objEndereco.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 208214

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208215

    'Sem Dados
    If lErro <> AD_SQL_SEM_DADOS Then
    
        objEnderecoBol.sSkype = sSkype
        objEnderecoBol.sRadio = sRadio
    
        Set objEndereco.objInfoUsu = objEnderecoBol
        
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Endereco_Le_Customizado = SUCESSO

    Exit Function

Erro_Endereco_Le_Customizado:

    Endereco_Le_Customizado = gErr

    Select Case gErr

        Case 208213
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208214, 208215
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ENDERECOS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208216)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function
