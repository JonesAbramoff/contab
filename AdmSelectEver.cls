VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Relatorio_Customiza_ExecutarAux(ByVal objRelOpcoes As AdmRelOpcoes, sNomeTsk As String) As Long

Dim lErro As Long
Dim sSerie As String


On Error GoTo Erro_Relatorio_Customiza_ExecutarAux

    'se forem relatorios de impressao de nota fiscal
    If UCase(left(sNomeTsk, 2)) = "NF" Then
        
        lErro = objRelOpcoes.ObterParametro("TSERIE", sSerie)
        If lErro <> SUCESSO Then gError 133003
        
        Select Case sSerie
        
            Case "1", "2", "3", "4", "5", "6", "7", "8", "1-e", "901-e"
        
                lErro = objRelOpcoes.IncluirParametro("NFILIAL", CStr(FilialEmpresa_ConvFRFO1(giFilialEmpresa)))
                If lErro <> AD_BOOL_TRUE Then gError 126936
            
                lErro = objRelOpcoes.IncluirParametro("TSERIE", "1")
                If lErro <> AD_BOOL_TRUE Then gError 133005
        
            Case Else
        
                lErro = objRelOpcoes.IncluirParametro("NFILIAL", CStr(giFilialEmpresa))
                If lErro <> AD_BOOL_TRUE Then gError 133004
        
        End Select
        
    ElseIf UCase(left(sNomeTsk, 5)) = "DANFE" Then

        lErro = objRelOpcoes.IncluirParametro("NFILIAL", CStr(FilialEmpresa_ConvFRFO1(giFilialEmpresa)))
        If lErro <> AD_BOOL_TRUE Then gError 126936
    
        lErro = objRelOpcoes.IncluirParametro("TSERIE", "1-e")
        If lErro <> AD_BOOL_TRUE Then gError 133005
    Else

        lErro = objRelOpcoes.IncluirParametro("NFILIAL", CStr(giFilialEmpresa))
        If lErro <> AD_BOOL_TRUE Then gError 126935

    End If

    Relatorio_Customiza_ExecutarAux = SUCESSO
    
    Exit Function
    
Erro_Relatorio_Customiza_ExecutarAux:

    Relatorio_Customiza_ExecutarAux = gErr
    
    Select Case gErr
    
        Case 126935, 126936, 133003, 133004, 133005
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142614)
        
    End Select
    
    Exit Function

End Function

Private Function FilialEmpresa_ConvFRFO1(ByVal iFilialEmpresa As Integer) As Integer

    If iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL1 Then
        FilialEmpresa_ConvFRFO1 = iFilialEmpresa
    Else
        FilialEmpresa_ConvFRFO1 = iFilialEmpresa - DELTA_FILIALREAL_OFICIAL1
    End If
    
End Function
