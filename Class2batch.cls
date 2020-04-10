VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class2batch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function Comando_BindVarInt1 Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt1 Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt1 Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Function Apuracao_Exercicio_Executa_SQL(ByVal sSQL As String, vConta1 As Variant, avConta() As Variant, ByVal lComando As Long, ByVal colContasApuracao As Collection) As Long
'prepara o comando SQL passado como parametro

Dim lErro As Long
Dim iIndice As Integer
Dim vConta As Variant
Dim vTipoConta As Variant
Dim sContaAux As String

On Error GoTo Erro_Apuracao_Exercicio_Executa_SQL

    lErro = Comando_PrepararInt1(lComando, sSQL)
    If lErro <> AD_SQL_SUCESSO Then Error 9777
    
    lErro = Comando_BindVarInt1(lComando, vConta1)
    If lErro <> AD_SQL_SUCESSO Then Error 9779
    
    vTipoConta = CONTA_ANALITICA
    
    lErro = Comando_BindVarInt1(lComando, vTipoConta)
    If lErro <> AD_SQL_SUCESSO Then Error 9780
    
    iIndice = 0
    
    For Each vConta In colContasApuracao

        avConta(iIndice) = vConta
        sContaAux = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaUltimaConta(avConta(iIndice), sContaAux)
        If lErro <> SUCESSO Then Error 9778

        avConta(iIndice + 1) = sContaAux
        lErro = Comando_BindVarInt1(lComando, avConta(iIndice))
        If lErro <> AD_SQL_SUCESSO Then Error 9779
            
        lErro = Comando_BindVarInt1(lComando, avConta(iIndice + 1))
        If lErro <> AD_SQL_SUCESSO Then Error 9780
        
        iIndice = iIndice + 2
        
    Next
    
    lErro = Comando_ExecutarInt1(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9781
    
    Apuracao_Exercicio_Executa_SQL = SUCESSO
    
    Exit Function

Erro_Apuracao_Exercicio_Executa_SQL:

    Apuracao_Exercicio_Executa_SQL = Err
    
    Select Case Err
    
        Case 9777, 9779, 9780, 9781
            
        Case 9778
        
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144629)
        
    End Select

    Exit Function
        
End Function

Function Apuracao_Exercicio_Comando_SQL(sSQL As String, ByVal colContasApuracao As Collection) As Long

Dim lErro As Long
Dim vConta As Variant

On Error GoTo Erro_Apuracao_Exercicio_Comando_SQL

    sSQL = "SELECT Conta FROM PlanoConta WHERE TipoConta = ? AND ("

    For Each vConta In colContasApuracao

        sSQL = sSQL & "(Conta >= ? AND Conta <= ?) OR "
        
    Next
    
    sSQL = Left(sSQL, Len(sSQL) - 4)

    sSQL = sSQL & ")"

    Apuracao_Exercicio_Comando_SQL = SUCESSO

    Exit Function

Erro_Apuracao_Exercicio_Comando_SQL:

    Apuracao_Exercicio_Comando_SQL = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144630)

    End Select

    Exit Function

End Function



