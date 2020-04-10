VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Dim sCodigoOSAnt As Long

Dim gcolMovEst As Collection

Public Function MovEstSRV_ObterCodProd(sTexto As String, ByVal iFilialEmpresa As Integer, ByVal sCodigoOS As String, ByVal iLinha As Integer) As Long
'Devolve o Código do Produto

Dim lErro As Long
Dim objItemMovEst As ClassItemMovEstoque

On Error GoTo Erro_MovEstSRV_ObterCodProd

    lErro = ObterDadosMovEstSRV(iFilialEmpresa, sCodigoOS)
    If lErro <> SUCESSO Then gError 210547
    
    If iLinha > gcolMovEst.Count Then
        Set objItemMovEst = Nothing
        sTexto = ""
    Else
        Set objItemMovEst = gcolMovEst(iLinha)
        sTexto = objItemMovEst.sProduto
    End If
    
    MovEstSRV_ObterCodProd = SUCESSO
    
    Exit Function

Erro_MovEstSRV_ObterCodProd:

    MovEstSRV_ObterCodProd = gErr
    
    Select Case gErr
        
        Case 210547
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210548)

    End Select

    Exit Function

End Function

Public Function MovEstSRV_ObterDescProd(sTexto As String, ByVal iFilialEmpresa As Integer, ByVal sCodigoOS As String, ByVal iLinha As Integer) As Long
'Devolve a descricao do Produto

Dim lErro As Long
Dim objItemMovEst As ClassItemMovEstoque

On Error GoTo Erro_MovEstSRV_ObterDescProd

    lErro = ObterDadosMovEstSRV(iFilialEmpresa, sCodigoOS)
    If lErro <> SUCESSO Then gError 210549
    
    If iLinha > gcolMovEst.Count Then
        Set objItemMovEst = Nothing
        sTexto = ""
    Else
        Set objItemMovEst = gcolMovEst(iLinha)
        sTexto = objItemMovEst.sDescricao
    End If
    
    MovEstSRV_ObterDescProd = SUCESSO
    
    Exit Function

Erro_MovEstSRV_ObterDescProd:

    MovEstSRV_ObterDescProd = gErr
    
    Select Case gErr
        
        Case 210549
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210550)

    End Select

    Exit Function

End Function

Public Function MovEstSRV_ObterSiglaUMProd(sTexto As String, ByVal iFilialEmpresa As Integer, ByVal sCodigoOS As String, ByVal iLinha As Integer) As Long
'Devolve a UM do Produto

Dim lErro As Long
Dim objItemMovEst As ClassItemMovEstoque

On Error GoTo Erro_MovEstSRV_ObterSiglaUMProd

    lErro = ObterDadosMovEstSRV(iFilialEmpresa, sCodigoOS)
    If lErro <> SUCESSO Then gError 210551
    
    If iLinha > gcolMovEst.Count Then
        Set objItemMovEst = Nothing
        sTexto = ""
    Else
        Set objItemMovEst = gcolMovEst(iLinha)
        sTexto = objItemMovEst.sSiglaUM
    End If
    
    MovEstSRV_ObterSiglaUMProd = SUCESSO
    
    Exit Function

Erro_MovEstSRV_ObterSiglaUMProd:

    MovEstSRV_ObterSiglaUMProd = gErr
    
    Select Case gErr
        
        Case 210551
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210552)

    End Select

    Exit Function

End Function

Public Function MovEstSRV_ObterQuantidade(dQtde As Double, ByVal iFilialEmpresa As Integer, ByVal sCodigoOS As String, ByVal iLinha As Integer) As Long
'Devolve a Quantidade do Produto

Dim lErro As Long
Dim objItemMovEst As ClassItemMovEstoque

On Error GoTo Erro_MovEstSRV_ObterQuantidade

    lErro = ObterDadosMovEstSRV(iFilialEmpresa, sCodigoOS)
    If lErro <> SUCESSO Then gError 210553
    
    If iLinha > gcolMovEst.Count Then
        Set objItemMovEst = Nothing
        dQtde = 0
    Else
        Set objItemMovEst = gcolMovEst(iLinha)
        dQtde = objItemMovEst.dQuantidade
    End If
    
    MovEstSRV_ObterQuantidade = SUCESSO
    
    Exit Function

Erro_MovEstSRV_ObterQuantidade:

    MovEstSRV_ObterQuantidade = gErr
    
    Select Case gErr
        
        Case 210553
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210554)

    End Select

    Exit Function

End Function

Private Function ObterDadosMovEstSRV(ByVal iFilialEmpresa As Integer, ByVal sCodigoOS As String) As Long
'Verifica se está realizando a leitura de uma OS Diferente

Dim lErro As Long

On Error GoTo Erro_ObterDadosMovEstSRV
 
    'se o Codigo da OS passado como parametro mudou
    If sCodigoOS <> sCodigoOSAnt Then
               
        Set gcolMovEst = New Collection
               
        'le as movimentecoes de estoque de saida de material para conserto associadas a OS
        lErro = CF("MovEstoque_Le_MovSRV", iFilialEmpresa, sCodigoOS, gcolMovEst)
        If lErro <> SUCESSO Then gError 210555
                        
        'seta como corrente o novo sCodigoOS
        sCodigoOSAnt = sCodigoOS
        
    End If
    
    ObterDadosMovEstSRV = SUCESSO

    Exit Function

Erro_ObterDadosMovEstSRV:

    ObterDadosMovEstSRV = gErr

    Select Case gErr
        
        Case 210555
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210556)

    End Select

    Exit Function

End Function

