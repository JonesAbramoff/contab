VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassClienteNome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable to hold collection
Private mvaraobjCliente() As ClassCliente
Private mvarlNumItens As Long
Private mvarlNumItensAlocados As Long

Public Function Add(objCliente As ClassCliente) As ClassCliente
    
Dim objNovo As ClassCliente
Dim objAntigo As ClassCliente
Dim lIndice As Long
Dim lIndice2 As Long

    Set objNovo = New ClassCliente
    
    'copiar atributos
    objNovo.lCodigo = objCliente.lCodigo
    objNovo.sNomeReduzido = objCliente.sNomeReduzido
    objNovo.sCGC = objCliente.sCGC
    
    'insere no array
    
    If (mvarlNumItensAlocados <= mvarlNumItens) Or mvarlNumItensAlocados = 0 Then
    
        mvarlNumItensAlocados = mvarlNumItensAlocados + 100
    
        ReDim Preserve mvaraobjCliente(mvarlNumItensAlocados)
    
    End If

    If mvarlNumItens > 0 Then
        Set objAntigo = Busca(objNovo.sNomeReduzido, lIndice)
        mvarlNumItens = mvarlNumItens + 1
    Else
        mvarlNumItens = mvarlNumItens + 1
        lIndice = mvarlNumItens
    End If
    
        
    For lIndice2 = mvarlNumItens - 1 To lIndice Step -1
        Set mvaraobjCliente(lIndice2 + 1) = mvaraobjCliente(lIndice2)
    Next
 
    Set mvaraobjCliente(lIndice) = objNovo

    'return the object created
    Set Add = objNovo

End Function

Public Property Get Item(lIndice As Long) As ClassCliente
  Set Item = mvaraobjCliente(lIndice)
End Property

Public Property Get Count() As Long
'    Count = mvarlNumItens

Dim lCount As Long

     Call CF_ECF("Clientes_Le_Count", lCount)
     Count = lCount

End Property


Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
'    Set NewEnum = mCol.[_NewEnum]
End Property

Public Property Get Busca(sNome As String, Optional lIndice As Long) As ClassCliente
'Pesquisa binária

Dim lInicio As Long
Dim lFim As Long
Dim lMeio As Long
Dim bAchou As Boolean
Dim objCliente As New ClassCliente
Dim lErro As Long
        
    If mvarlNumItens > 0 Then
        
        lInicio = 1
        lFim = mvarlNumItens
        
        bAchou = False
    
        Do While lFim >= lInicio
    
            If lMeio = Arredonda_ParaCima((lInicio + lFim) / 2) Then Exit Do
    
            lMeio = Arredonda_ParaCima((lInicio + lFim) / 2)
    
            If UCase(mvaraobjCliente(lMeio).sNomeReduzido) > UCase(sNome) Then
               lFim = lMeio - 1
            Else
                If UCase(mvaraobjCliente(lMeio).sNomeReduzido) < UCase(sNome) Then
                    lInicio = lMeio + 1
                Else
                    bAchou = True
                    lInicio = lFim + 1
                End If
            End If
        Loop
        
        If lMeio > 0 Then
            If UCase(mvaraobjCliente(lMeio).sNomeReduzido) < UCase(sNome) Then
                lIndice = lMeio + 1
            Else
                lIndice = lMeio
            End If
        Else
            lIndice = lMeio
        End If
        
        
        If bAchou Then
            Set Busca = mvaraobjCliente(lMeio)
        Else
            Set Busca = Nothing
        End If
    
    Else
    
        lErro = CF_ECF("Clientes_Le_NomeReduzido", sNome, objCliente)
        If lErro <> SUCESSO Then
            Set Busca = Nothing
        Else
            Set Busca = objCliente
        End If
    
    End If
    
    
End Property

Private Sub Class_Initialize()
    mvarlNumItensAlocados = 0
End Sub


