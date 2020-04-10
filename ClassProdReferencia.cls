VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdReferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable to hold collection
Private mvaraobjProduto() As ClassProduto
Private mvarlNumItens As Long
Private mvarlNumItensAlocados As Long

Public Function Add(objProduto As ClassProduto) As ClassProduto
    
Dim objNovo As ClassProduto
Dim objAntigo As ClassProduto
Dim iIndice As Integer
Dim iIndice2 As Integer

    Set objNovo = New ClassProduto
    
    'copiar atributos
    objNovo.sCodigo = objProduto.sCodigo
    objNovo.sNomeReduzido = objProduto.sNomeReduzido
    objNovo.sSiglaUMVenda = objProduto.sSiglaUMVenda
    objNovo.sReferencia = objProduto.sReferencia
    objNovo.sFigura = objProduto.sFigura
    objNovo.dPrecoLoja = objProduto.dPrecoLoja
    objNovo.sSituacaoTribECF = objProduto.sSituacaoTribECF
    objNovo.sICMSAliquota = objProduto.sICMSAliquota
    objNovo.dPercentMenosReceb = objProduto.dPercentMenosReceb
    objNovo.sDescricao = objProduto.sDescricao
    objNovo.dDescontoValor = objProduto.dDescontoValor
    objNovo.iUsaBalanca = objProduto.iUsaBalanca
    Set objNovo.colCodBarras = objProduto.colCodBarras
    objNovo.iCompras = objProduto.iCompras
    objNovo.sTruncamento = objProduto.sTruncamento
    objNovo.dQuantEstLoja = objProduto.dQuantEstLoja
    Set objNovo.colCodBarras = objProduto.colCodBarras
    
    '??? em vez de ficar copiando todos os campos poderia referenciar o proprio objProduto (Jones 2603/14)
    Call objNovo.objTributacaoDocItem.Copia(objProduto.objTributacaoDocItem)
    
    'insere no array
     
    If (mvarlNumItensAlocados < mvarlNumItens + 1) Or mvarlNumItensAlocados = 0 Then
    
        mvarlNumItensAlocados = mvarlNumItensAlocados + 100
    
        ReDim Preserve mvaraobjProduto(mvarlNumItensAlocados)
    
    End If

    
    If mvarlNumItens > 0 Then
        Set objAntigo = Busca(objNovo.sReferencia, iIndice)
        mvarlNumItens = mvarlNumItens + 1
    Else
        mvarlNumItens = mvarlNumItens + 1
        iIndice = mvarlNumItens
    End If
    
        
    For iIndice2 = mvarlNumItens - 1 To iIndice Step -1
        Set mvaraobjProduto(iIndice2 + 1) = mvaraobjProduto(iIndice2)
    Next
 
    Set mvaraobjProduto(iIndice) = objNovo
    
    'return the object created
    Set Add = objNovo

    Set objNovo = Nothing
    
End Function

Public Property Get Item(iIndice As Integer) As ClassProduto
  Set Item = mvaraobjProduto(iIndice)
End Property

Public Property Get Items(sReferencia As String) As ClassProduto
  Set Items = mvaraobjProduto(sReferencia)
End Property

Public Property Get Count() As Long
    Count = mvarlNumItens
End Property

Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
'    Set NewEnum = mCol.[_NewEnum]
End Property

Public Property Get Busca(sReferencia As String, Optional iIndice As Integer) As ClassProduto
'Pesquisa binária

Dim iInicio As Integer
Dim iFim As Integer
Dim iMeio As Integer
Dim bAchou As Boolean
Dim objProduto As New ClassProduto
Dim lErro As Long
        
    If mvarlNumItens > 0 Then
        
        iInicio = 1
        iFim = mvarlNumItens
        
        bAchou = False
    
        Do While iFim >= iInicio
    
            If iMeio = Arredonda_ParaCima((iInicio + iFim) / 2) Then Exit Do
    
            iMeio = Arredonda_ParaCima((iInicio + iFim) / 2)
    
            If UCase(mvaraobjProduto(iMeio).sReferencia) > UCase(sReferencia) Then
               iFim = iMeio - 1
            Else
                If UCase(mvaraobjProduto(iMeio).sReferencia) < UCase(sReferencia) Then
                    iInicio = iMeio + 1
                Else
                    bAchou = True
                    iInicio = iFim + 1
                End If
            End If
        Loop
        
        If iMeio > 0 Then
            If UCase(mvaraobjProduto(iMeio).sReferencia) < UCase(sReferencia) Then
                iIndice = iMeio + 1
            Else
                iIndice = iMeio
            End If
        Else
            iIndice = iMeio
        End If
        
        If bAchou Then
            
            Set Busca = mvaraobjProduto(iMeio)
        
        Else
            
            Set Busca = Nothing
        
        End If
    
    Else
    
        lErro = CF_ECF("Produtos_Le_Referencia", sReferencia, objProduto)
        If lErro <> SUCESSO Then
            Set Busca = Nothing
        Else
            Set Busca = objProduto
        End If
    
    End If
    
End Property

Public Property Get BuscaNome(sNome As String) As ClassProduto
'Pesquisa binária

Dim iInicio As Integer
Dim iFim As Integer
Dim iMeio As Integer
Dim bAchou As Boolean
        
    iInicio = 1
    iFim = mvarlNumItens
    
    bAchou = False

    Do While iFim >= iInicio

        iMeio = (iInicio + iFim) / 2

        If UCase(mvaraobjProduto(iMeio).sNomeReduzido) > UCase(sNome) Then
           iFim = iMeio - 1
        Else
            If UCase(mvaraobjProduto(iMeio).sNomeReduzido) < UCase(sNome) Then
                iInicio = iMeio + 1
            Else
                bAchou = True
                iInicio = iFim + 1
            End If
        End If
    Loop
    
    If bAchou Then
        Set BuscaNome = mvaraobjProduto(iMeio)
    Else
        Set BuscaNome = Nothing
    End If
    
End Property

Private Sub Class_Initialize()
    mvarlNumItensAlocados = 0
End Sub
