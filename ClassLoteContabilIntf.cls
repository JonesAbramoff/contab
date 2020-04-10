VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLoteContabilIntf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjLote As Object

Public Property Set objLote(ByVal vData As Object)
    Set mvarobjLote = vData
End Property

Public Property Get objLote() As Object
    Set objLote = mvarobjLote
End Property

Public Function IniciarDoc(ByVal dtData As Date, objDocContabilIntf As ClassDocContabilIntf) As Long
'recebe a Data e retorna o voucher

Dim lErro As Long
Dim lDoc As Long
Dim objPeriodo As New ClassPeriodo

On Error GoTo Erro_IniciarDoc

    lErro = CF("Periodo_Le", dtData, objPeriodo)
    If lErro <> SUCESSO Then gError 92930

    If objPeriodo.iExercicio <> mvarobjLote.iExercicio Then gError 92931

    If objPeriodo.iPeriodo <> mvarobjLote.iPeriodo Then gError 92932

    lErro = CF("Voucher_Automatico", objLote.iFilialEmpresa, objLote.iExercicio, objLote.iPeriodo, objLote.sOrigem, lDoc)
    If lErro <> SUCESSO Then gError 92933
    
    Set objDocContabilIntf = New ClassDocContabilIntf
    
    objDocContabilIntf.dtData = dtData
    objDocContabilIntf.lDoc = lDoc
    Set objDocContabilIntf.objLote = mvarobjLote

    IniciarDoc = SUCESSO
    
    Exit Function

Erro_IniciarDoc:

    IniciarDoc = gErr

    Select Case gErr

        Case 92930, 92933

        Case 92931
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIOS_DIFERENTES", gErr, objPeriodo.iExercicio, objLote.iExercicio)

        Case 92932
            Call Rotina_Erro(vbOKOnly, "ERRO_PERIODOS_DIFERENTES", gErr, objPeriodo.iPeriodo, objLote.iPeriodo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151537)

    End Select
    
    Exit Function

End Function




