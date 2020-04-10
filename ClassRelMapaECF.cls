VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMapaECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Criada para atender a mapa de cotacao

Dim lComando As Long

Dim iCodECFAnt As Integer
Dim dtDataAnt As Date

Dim iUltimoNumerico As Integer

Dim colAliquotas As New Collection
Dim colLojaArqFisAnalitico As New Collection

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)
        
End Sub
'------------------------------------------------------------------------

Public Function MapaECF_ObterAliquota(dAliquota As Double, ByVal dtData As Date, ByVal iPOS As Integer) As Long

Dim lErro As Long
Dim objLojaArqFisAnalitico As ClassLojaArqFisAnalitico

On Error GoTo Erro_MapaECF_ObterAliquota

    lErro = ObterAliquotasDia(dtData)
    If lErro <> SUCESSO Then gError 140610
    
    If iPOS > iUltimoNumerico Then
        Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
    Else
        Set objLojaArqFisAnalitico = colAliquotas.Item(iPOS)
    End If
    
    dAliquota = StrParaDbl(objLojaArqFisAnalitico.sSituacaoTrib) / 100
    
    MapaECF_ObterAliquota = SUCESSO
    
    Exit Function

Erro_MapaECF_ObterAliquota:

    MapaECF_ObterAliquota = gErr
    
    Select Case gErr
    
        Case 140610
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153223)

    End Select

    Exit Function

End Function

Public Function MapaECF_ObterValorAliquota(dValor As Double, ByVal dtData As Date, ByVal iCodECF As Integer, ByVal sSituacaoTrib As String) As Long

Dim lErro As Long
Dim bAchou As Boolean
Dim objLojaArqFisAnalitico As ClassLojaArqFisAnalitico

On Error GoTo Erro_MapaECF_ObterValorAliquota

    lErro = ObterDadosMapaECF(dtData, iCodECF)
    If lErro <> SUCESSO Then gError 140615
    
    For Each objLojaArqFisAnalitico In colLojaArqFisAnalitico
    
        If Not IsNumeric(sSituacaoTrib) Then
            If UCase(Trim(objLojaArqFisAnalitico.sSituacaoTrib)) = UCase(Trim(sSituacaoTrib)) Then
                bAchou = True
                Exit For
            End If
        Else
            If IsNumeric(objLojaArqFisAnalitico.sSituacaoTrib) Then
                If Abs(StrParaDbl(objLojaArqFisAnalitico.sSituacaoTrib) / 100 - StrParaDbl(sSituacaoTrib)) < QTDE_ESTOQUE_DELTA Then
                    bAchou = True
                    Exit For
                End If
            End If
        End If
    
    Next
        
    If bAchou Then
        dValor = objLojaArqFisAnalitico.dTotalizador
    Else
        dValor = 0
    End If
    
    MapaECF_ObterValorAliquota = SUCESSO
    
    Exit Function

Erro_MapaECF_ObterValorAliquota:

    MapaECF_ObterValorAliquota = gErr
    
    Select Case gErr
    
        Case 140615
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153224)

    End Select

    Exit Function

End Function

Public Function ObterDadosMapaECF(ByVal dtData As Date, ByVal iCodECF As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_ObterDadosMapaECF
 
    If dtData <> dtDataAnt Or iCodECF <> iCodECFAnt Then
    
        Set colLojaArqFisAnalitico = New Collection
        
        lErro = LojaArqFisAnalitico_Le(dtData, iCodECF, colLojaArqFisAnalitico)
        If lErro <> SUCESSO Then gError 140616
        
        dtDataAnt = dtData
    
    End If
        
    ObterDadosMapaECF = SUCESSO

    Exit Function

Erro_ObterDadosMapaECF:

    ObterDadosMapaECF = gErr

    Select Case gErr
    
        Case 140616
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153225)

    End Select

    Exit Function

End Function

Public Function ObterAliquotasDia(ByVal dtData As Date) As Long

Dim lErro As Long

On Error GoTo Erro_ObterAliquotasDia
 
    If dtData <> dtDataAnt Then
    
        Set colAliquotas = New Collection
        
        lErro = LojaArqFisAnalitico_Le_AliquotasDia(dtData, colAliquotas, iUltimoNumerico)
        If lErro <> SUCESSO Then gError 140611
        
        dtDataAnt = dtData
    
    End If
        
    ObterAliquotasDia = SUCESSO

    Exit Function

Erro_ObterAliquotasDia:

    ObterAliquotasDia = gErr

    Select Case gErr
    
        Case 140611
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153226)

    End Select

    Exit Function

End Function

Function LojaArqFisAnalitico_Le_AliquotasDia(ByVal dtData As Date, ByVal colAliquotas As Collection, iUltimoNumerico As Integer) As Long

Dim lErro As Long
Dim sSituacaoTrib As String
Dim objLojaArqFisAnalitico As ClassLojaArqFisAnalitico

On Error GoTo Erro_LojaArqFisAnalitico_Le_AliquotasDia

    'Abre o comando
    If lComando = 0 Then
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 140617
    End If
    
    iUltimoNumerico = 0
    
    sSituacaoTrib = String(4, 0)
    
    lErro = Comando_Executar(lComando, "SELECT SituacaoTrib FROM LojaArqFisAnalitico WHERE Data = ? AND FilialEmpresa = ? GROUP BY SituacaoTrib ORDER BY SituacaoTrib", sSituacaoTrib, dtData, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 140618
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140619
    
    'Se nao encontrou => erro
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
    
        objLojaArqFisAnalitico.dtData = dtData
        objLojaArqFisAnalitico.iFilialEmpresa = giFilialEmpresa
        objLojaArqFisAnalitico.sSituacaoTrib = sSituacaoTrib
        
        colAliquotas.Add objLojaArqFisAnalitico
        
        If IsNumeric(objLojaArqFisAnalitico.sSituacaoTrib) Then iUltimoNumerico = iUltimoNumerico + 1

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140620

    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    LojaArqFisAnalitico_Le_AliquotasDia = SUCESSO

    Exit Function

Erro_LojaArqFisAnalitico_Le_AliquotasDia:

    LojaArqFisAnalitico_Le_AliquotasDia = gErr

    Select Case gErr

        Case 140617
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 140618 To 140620
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOJAARQFISANALITICO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153227)

    End Select
    
    Exit Function

End Function

Function LojaArqFisAnalitico_Le(ByVal dtData As Date, ByVal iCodECF As Integer, ByVal colLojaArqFisMestre As Collection) As Long

Dim lErro As Long
Dim sSituacaoTrib As String
Dim dTotalizador As Double
Dim objLojaArqFisAnalitico As ClassLojaArqFisAnalitico

On Error GoTo Erro_LojaArqFisAnalitico_Le

    'Abre o comando
    If lComando = 0 Then
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 140621
    End If
    
    sSituacaoTrib = String(4, 0)
    
    lErro = Comando_Executar(lComando, "SELECT SituacaoTrib, Totalizador FROM LojaArqFisAnalitico WHERE Data = ? AND FilialEmpresa = ? AND CodECF = ? ORDER BY SituacaoTrib", sSituacaoTrib, dTotalizador, dtData, giFilialEmpresa, iCodECF)
    If lErro <> AD_SQL_SUCESSO Then gError 140622
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140623
    
    'Se nao encontrou => erro
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
    
        objLojaArqFisAnalitico.dtData = dtData
        objLojaArqFisAnalitico.dTotalizador = dTotalizador
        objLojaArqFisAnalitico.iCodECF = iCodECF
        objLojaArqFisAnalitico.iFilialEmpresa = giFilialEmpresa
        objLojaArqFisAnalitico.sSituacaoTrib = sSituacaoTrib
        
        colLojaArqFisAnalitico.Add objLojaArqFisAnalitico

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140624

    Loop
    
    LojaArqFisAnalitico_Le = SUCESSO

    Exit Function

Erro_LojaArqFisAnalitico_Le:

    LojaArqFisAnalitico_Le = gErr

    Select Case gErr

        Case 140621
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 140622 To 140624
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOJAARQFISANALITICO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153228)

    End Select

    Exit Function

End Function
