VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFolhaIntf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjIntfCtb As New ClassIntfCtb

Public Property Get objIntfCtb() As Object
    Set objIntfCtb = mvarobjIntfCtb
End Property

Public Function Login(ByVal sEmpresa As String, ByVal sFilial As String) As Long

Dim lErro As Long, objUsuario As New ClassDicUsuario
Dim objEmpresa As New ClassDicEmpresa
Dim objFilialEmpresa As New ClassFilialEmpresa
Dim colFilialEmpresa As New Collection
Dim objUsuarioEmpresa As ClassUsuarioEmpresa
Dim iAchou As Integer
Dim lSistema As Long
Dim colFiliais As New Collection
Dim colModulo As New AdmColModulo
Dim sUsuario As String
Dim lCodigo As Long

On Error GoTo Erro_Login

    'para permitir acessar o dicionario de dados
    lSistema = Sistema_Abrir()
    If lSistema = 0 Then gError 92935

    'Preenche a chave de objUsuarios
    objUsuario.sCodUsuario = "supervisor"

    lErro = DicUsuario_Le(objUsuario)
    If lErro <> SUCESSO Then gError 92936
    
    'faz login utilizando o codigo do usuario e a senha
    lErro = Usuario_Executa_Login(objUsuario.sCodUsuario, objUsuario.sSenha)
    If lErro <> SUCESSO Then gError 92937
    
    objEmpresa.sNome = sEmpresa
    
    lErro = Empresa_Le_Nome(objEmpresa)
    If lErro <> SUCESSO And lErro <> 6289 Then gError 92938

    If lErro = 6289 Then gError 92910

    sUsuario = "supervisor"
    lCodigo = objEmpresa.lCodigo

    'Carregar todas as filiais da empresa selecionada para os quais o usuário está autorizado a acessar
    lErro = FiliaisEmpresa_Le_Usuario(sUsuario, lCodigo, colFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 50172 Then gError 92940

    For Each objUsuarioEmpresa In colFilialEmpresa
        
        If objUsuarioEmpresa.sNomeFilial = sFilial Then
            iAchou = 1
            Exit For
        End If
        
    Next
    
    If iAchou <> 1 Then gError 92915
    
    'seleciona a Empresa e filial
    lErro = Sistema_DefEmpresa(sEmpresa, objEmpresa.lCodigo, sFilial, objUsuarioEmpresa.iCodFilial)
    If lErro <> AD_BOOL_TRUE Then gError 92939
    
    'Carrega em gcolModulo os módulos indicando ativadade p/ FilialEmpresa
    lErro = CF("Modulos_Le_Empresa_Filial", objEmpresa.lCodigo, objUsuarioEmpresa.iCodFilial, colModulo)
    If lErro <> SUCESSO Then gError 92913

    If colModulo.Ativo(MODULO_CONTABILIDADE) <> MODULO_ATIVO Then gError 92914

    gdtDataHoje = Date
    
    gdtDataAtual = gdtDataHoje

    'Ler o Codigo da Empresa
    mvarobjIntfCtb.lCodEmpresa = objEmpresa.lCodigo
    mvarobjIntfCtb.sUsuario = "supervisor"
    mvarobjIntfCtb.iFilialEmpresa = objUsuarioEmpresa.iCodFilial

    Login = SUCESSO
    
    Exit Function

Erro_Login:

    Login = gErr

    Select Case gErr

        Case 92910
            Call Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_NAO_CADASTRADA1", gErr, objEmpresa.sNome)

        Case 92913, 92935, 92936, 92937, 92938, 92939, 92940

        Case 92914
            Call Rotina_Erro(vbOKOnly, "ERRO_MODULO_CTB_INATIVO", gErr)

        Case 92915
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_SEM_FILIAL", gErr, "supervisor", sFilial)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150473)

    End Select
    
    Exit Function

End Function

Public Function Logoff() As Long

Dim lErro As Long
Dim objEmpresa As New ClassDicEmpresa
Dim lCodEmpresa As Long
Dim colFiliais As New Collection

On Error GoTo Erro_Logoff

    Call Usuario_Altera_SituacaoLogin(gsUsuario, USUARIO_NAO_LOGADO)
    
    Call Sistema_Fechar
    
    Logoff = SUCESSO
   
    Exit Function

Erro_Logoff:

    Logoff = gErr

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150474)

    End Select
    
    Exit Function

End Function

