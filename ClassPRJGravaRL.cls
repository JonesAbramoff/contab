VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJGravaRL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Projetos_Grava_Customizado(ByVal objProjetos As ClassProjetos) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lDemonstrativo As Long
Dim lDemonstrativoAnt As Long

On Error GoTo Erro_Projetos_Grava_Customizado

    If objProjetos.dtDataFim <> DATA_NULA Then
    
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 181313
        Next
    
        'Le a tabelaProjetos
        lErro = Comando_Executar(alComando(2), "SELECT MAX(Demonstrativo) FROM Projetos WHERE YEAR(DataCriacao) = ? ", _
                     lDemonstrativo, Year(objProjetos.dtDataCriacao))
        If lErro <> AD_SQL_SUCESSO Then gError 181314
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181315
        
        lDemonstrativo = lDemonstrativo + 1
    
        'Le a tabelaProjetos
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Demonstrativo FROM Projetos WHERE NumIntDoc = ? ", _
                    0, lDemonstrativoAnt, objProjetos.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 181314
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181315
        
        'Se existir => UPDATE
        If lErro = AD_SQL_SUCESSO And lDemonstrativoAnt = 0 Then
    
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(0))
            If lErro <> SUCESSO Then gError 181316
    
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE Projetos SET Demonstrativo = ? ", alComando(0), _
                    lDemonstrativo)
            If lErro <> AD_SQL_SUCESSO Then gError 181317
    
        End If

        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    End If
    
    Projetos_Grava_Customizado = SUCESSO

    Exit Function

Erro_Projetos_Grava_Customizado:

    Projetos_Grava_Customizado = gErr

    Select Case gErr

        Case 181312, 181464, 181569

        Case 181313
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181314, 181315
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROJETOS", gErr)

        Case 181316
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PROJETOS", gErr)

        Case 181317
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PROJETOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181410)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function





