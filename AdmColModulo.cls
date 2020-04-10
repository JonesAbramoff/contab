VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmColModulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mvarcolModulo As New Collection

Public Property Get Item(vntIndexKey As Integer) As AdmModulo
    Set Item = mvarcolModulo(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mvarcolModulo.Count
End Property

Public Function Ativo(sSigla As String) As Long
'recebe a sigla do modulo e retorna a indicacao se o modulo está ativo ou inativo
'retorna -1 se não achou um código para a sigla fornecida

Dim objModulo As AdmModulo
Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objModulo In mvarcolModulo
    
        If objModulo.sSigla = sSigla Then
            Ativo = objModulo.iAtivo
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        Ativo = -1
    End If
    
End Function

Public Function Sigla(sNome As String) As String
'recebe o nome do modulo e retorna a sigla
'retorna "" se não achou uma sigla para o nome fornecido

Dim objModulo As AdmModulo
    
    Sigla = ""
    
    For Each objModulo In mvarcolModulo
    
        If objModulo.sNome = sNome Then
            Sigla = objModulo.sSigla
            Exit Function
        End If
    Next
    
End Function

Public Function Nome(sSigla As String) As String
'recebe a sigla do modulo e retorna o nome
'retorna "" se não achou um nome para a sigla fornecida

Dim objModulo As AdmModulo
    
    Nome = ""
    
    For Each objModulo In mvarcolModulo
    
        If objModulo.sSigla = sSigla Then
            Nome = objModulo.sNome
            Exit Function
        End If
    Next
    
End Function

Public Function sOrigemCusto(sSigla As String) As String
'recebe a sigla do modulo e retorna a origem dos lançamentos de custo associada ao modulo
'retorna "" se não achou um nome para a sigla fornecida

Dim objModulo As AdmModulo
    
    sOrigemCusto = ""
    
    For Each objModulo In mvarcolModulo
    
        If objModulo.sSigla = sSigla Then
            sOrigemCusto = objModulo.sOrigemCusto
            Exit Function
        End If
    Next
    
End Function

Public Function sOrigemAglutina(sSigla As String) As String
'recebe a sigla do modulo e retorna a origem dos lançamentos de aglutinação associada ao modulo
'retorna "" se não achou um nome para a sigla fornecida

Dim objModulo As AdmModulo
    
    sOrigemAglutina = ""
    
    For Each objModulo In mvarcolModulo
    
        If objModulo.sSigla = sSigla Or objModulo.sOrigemCusto = sSigla Then
            sOrigemAglutina = objModulo.sOrigemAglutina
            Exit Function
        End If
    Next
    
End Function

Public Function Add(sSigla As String, sNome As String, sDescricao As String, sVersao As String, sOrigemCusto As String, sOrigemAglutina As String, iAtivo As Integer, Optional sKey As String) As AdmModulo

Dim objNewMember As AdmModulo
Set objNewMember = New AdmModulo

    objNewMember.sSigla = sSigla
    objNewMember.sNome = sNome
    objNewMember.sDescricao = sDescricao
    objNewMember.sVersao = sVersao
    objNewMember.sOrigemCusto = sOrigemCusto
    objNewMember.sOrigemAglutina = sOrigemAglutina
    objNewMember.iAtivo = iAtivo

    If Len(sKey) = 0 Then
        mvarcolModulo.Add objNewMember
    Else
        mvarcolModulo.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

