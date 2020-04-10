VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIntfCtb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlCodEmpresa As Long
Private mvarsUsuario As String
Private mvariFilialEmpresa As Integer
Private mvariAno As Integer
Private mvariMes As Integer
Private mvariExercicio As Integer
Private mvariPeriodo As Integer

Public Property Let iPeriodo(ByVal vData As Integer)
    mvariPeriodo = vData
End Property

Public Property Get iPeriodo() As Integer
    iPeriodo = mvariPeriodo
End Property

Public Property Let iExercicio(ByVal vData As Integer)
    mvariExercicio = vData
End Property

Public Property Get iExercicio() As Integer
    iExercicio = mvariExercicio
End Property

Public Property Let iMes(ByVal vData As Integer)
    mvariMes = vData
End Property

Public Property Get iMes() As Integer
    iMes = mvariMes
End Property

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let lCodEmpresa(ByVal vData As Long)
    mvarlCodEmpresa = vData
End Property

Public Property Get lCodEmpresa() As Long
    lCodEmpresa = mvarlCodEmpresa
End Property

Public Function CriarLoteContabil(ByVal iAno As Integer, ByVal iMes As Integer, objLoteContabilIntf As ClassLoteContabilIntf) As Long
'recebe a Filial, Ano e Mes, cria e retorna o lote

Dim lErro As Long
Dim objEmpresa As New ClassDicEmpresa
Dim lCodEmpresa As Long
Dim iAchou As Integer
Dim objUsuarioEmpresa As ClassUsuarioEmpresa
Dim objLote As New ClassLote
Dim objPeriodo As New ClassPeriodo
Dim dtData As Date
Dim objObject As Object

On Error GoTo Erro_CriarLoteContabil

    If Not IsDate("01/" & CStr(iMes) & "/" & CStr(iAno)) Then gError 92916
    
    dtData = CDate("01/" & CStr(iMes) & "/" & CStr(iAno))

    lErro = CF("Periodo_Le1", dtData, objPeriodo, mvariFilialEmpresa)
    If lErro <> SUCESSO Then gError 92917

    mvariAno = iAno
    mvariMes = iMes
    mvariExercicio = objPeriodo.iExercicio
    mvariPeriodo = objPeriodo.iPeriodo
    
    Set objObject = Me
    
    'Cria o lote contabil para a rotina de interação com a folha de pagamento
    lErro = CF("Lote_Contabil_Folha", objObject, objLote)
    If lErro <> SUCESSO Then gError 92922

    Set objLoteContabilIntf = New ClassLoteContabilIntf

    Set objLoteContabilIntf.objLote = objLote

    CriarLoteContabil = SUCESSO
    
    Exit Function

Erro_CriarLoteContabil:

    CriarLoteContabil = gErr

    Select Case gErr

        Case 92911, 92913, 92917, 92918, 92922

        Case 92916
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_MES_INVALIDO", gErr, iAno, iMes)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150896)

    End Select
    
    Exit Function

End Function

