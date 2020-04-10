VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMvPerCcl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim lComando As Long
Dim iFilialEmpresaAnt As Integer
Dim iExercicioAnt As Integer
Dim iPeriodoAnt As Integer
Dim sCclAnt As String
Dim dValorAnt As Double

Function ObterSldIniCcl(dValor As Double, iFilialEmpresa As Integer, iExercicio As Integer, iPeriodo As Integer, sCcl As String) As Long

Dim lErro As Long

On Error GoTo Erro_ObterSldIniCcl

    If iFilialEmpresa <> iFilialEmpresaAnt Or iExercicio <> iExercicioAnt Or iPeriodo <> iPeriodoAnt Or sCcl <> sCclAnt Then
            
        lErro = CF("MvPerCcl_Le_SldAnterior",iFilialEmpresa, iExercicio, sCcl, iPeriodo, dValor, lComando)
        If lErro <> SUCESSO Then gError 78083
        
        iFilialEmpresaAnt = iFilialEmpresa
        iExercicioAnt = iExercicio
        iPeriodoAnt = iPeriodo
        sCclAnt = sCcl
        dValorAnt = dValor
        
    End If
    
    dValor = dValorAnt
        
    Exit Function
    
    ObterSldIniCcl = SUCESSO
    
Erro_ObterSldIniCcl:

    ObterSldIniCcl = gErr
    
    Select Case gErr
        
        Case 78083
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153277)

    End Select

    Exit Function

End Function

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)
    
End Sub
