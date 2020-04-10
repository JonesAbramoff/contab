VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmFiliaisAdmCad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarsCodigo As String
Private mvarsInscricao As String

Public Property Let sInscricao(ByVal vData As String)
    mvarsInscricao = vData
End Property

Public Property Get sInscricao() As String
    sInscricao = mvarsInscricao
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Sub Copiar(objFilialOrig As AdmFiliaisAdmCad)
'atualiza o proprio conteudo copiando dados de outra instancia

    mvariFilialEmpresa = objFilialOrig.iFilialEmpresa
    mvarsCodigo = objFilialOrig.sCodigo
    mvarsInscricao = objFilialOrig.sInscricao

End Sub
