VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJSelectRL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Projetos_Le_Customizado(ByVal objProjetos As ClassProjetos) As Long

Dim lErro As Long
Dim alComando(0 To 0) As Long
Dim iIndice As Integer
Dim objProjetoRJ As New ClassProjetosRL
Dim lDemonstrativo As Long

On Error GoTo Erro_Projetos_Le_Customizado
   
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181313
    Next

    'Le a tabelaProjetos
    lErro = Comando_Executar(alComando(0), "SELECT Demonstrativo FROM Projetos WHERE NumIntDoc = ? ", _
                lDemonstrativo, objProjetos.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 181314

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181315
    
    objProjetoRJ.lDemonstrativo = lDemonstrativo
    Set objProjetos.objInfoUsu = objProjetoRJ
    
     'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Projetos_Le_Customizado = SUCESSO

    Exit Function

Erro_Projetos_Le_Customizado:

    Projetos_Le_Customizado = gErr

    Select Case gErr

        Case 181312, 181464, 181569

        Case 181313
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181314, 181315
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181410)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

