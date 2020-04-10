VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelFluxoCxMiguez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private dtUltData As Date
Private colRec As Collection
Private colPag As Collection

Private lComando As Long

Function FlCx_ObterValor(dValor As Double, ByVal lNumIntRel As Long, ByVal iTipo As Integer, ByVal dtData As Date, ByVal iLinha As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_FlCx_ObterValor

    lErro = ObterDadosFluxo(lNumIntRel, dtData)
    If lErro <> SUCESSO Then gError 33333
    
    If iTipo = NATUREZA_TIPO_RECEBIMENTO Then
    
        If iLinha >= 1 And iLinha <= colRec.Count Then
            dValor = colRec.Item(iLinha).dValor
        Else
            dValor = 0
        End If
        
    Else
        
        If iLinha >= 1 And iLinha <= colPag.Count Then
            dValor = colPag.Item(iLinha).dValor
        Else
            dValor = 0
        End If
    
    End If
    
    FlCx_ObterValor = SUCESSO
     
    Exit Function
    
Erro_FlCx_ObterValor:

    FlCx_ObterValor = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Function FlCx_ObterDescricao(sDescricao As String, ByVal lNumIntRel As Long, ByVal iTipo As Integer, ByVal dtData As Date, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objRelFluxoCxItem As ClassRelFluxoCxItem

On Error GoTo Erro_FlCx_ObterDescricao

    lErro = ObterDadosFluxo(lNumIntRel, dtData)
    If lErro <> SUCESSO Then gError 33333
    
    If iTipo = NATUREZA_TIPO_RECEBIMENTO Then
    
        If iLinha >= 1 And iLinha <= colRec.Count Then
            sDescricao = colRec.Item(iLinha).sDescricao
        Else
            sDescricao = ""
        End If
        
    Else
        
        If iLinha >= 1 And iLinha <= colPag.Count Then
            Set objRelFluxoCxItem = colPag.Item(iLinha)
            sDescricao = objRelFluxoCxItem.sDescricao
            If objRelFluxoCxItem.iNumParcelas > 1 Then sDescricao = sDescricao & " " & CStr(objRelFluxoCxItem.iNumParcela) & "/" & CStr(objRelFluxoCxItem.iNumParcelas)
        Else
            sDescricao = ""
        End If
    
    End If
    
    FlCx_ObterDescricao = SUCESSO
     
    Exit Function
    
Erro_FlCx_ObterDescricao:

    FlCx_ObterDescricao = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Function ObterDadosFluxo(ByVal lNumIntRel As Long, ByVal dtData As Date) As Long

Dim lErro As Long
Dim iSequencial As Integer, iTipo As Integer, iFluxoCaixa As Integer, sDescricao As String, dValor As Double, iNumParcelas As Integer, iNumParcela As Integer, iTipoNumIntDocOrigem As Integer, lNumIntDocOrigem As Long
Dim objRelFluxoCxItem As ClassRelFluxoCxItem

On Error GoTo Erro_ObterDadosFluxo

    If dtData <> dtUltData Then

        Set colRec = New Collection
        Set colPag = New Collection

        dtUltData = dtData

        If lComando = 0 Then

            lComando = Comando_AbrirExt(GL_lConexaoRel)
            If lComando = 0 Then gError 33333

        End If

        sDescricao = String(255, 0)
        
        lErro = Comando_Executar(lComando, "SELECT Sequencial, Tipo, FluxoCaixa, Descricao, Valor, NumParcelas, NumParcela, TipoNumIntDocOrigem, NumIntDocOrigem FROM RelOpFluxoCxItemMiguez WHERE NumIntRel = ? AND Data = ? ORDER BY Tipo, Valor DESC, Sequencial", _
            iSequencial, iTipo, iFluxoCaixa, sDescricao, dValor, iNumParcelas, iNumParcela, iTipoNumIntDocOrigem, lNumIntDocOrigem, lNumIntRel, dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 33333

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333

        Do While lErro = AD_SQL_SUCESSO
        
            Set objRelFluxoCxItem = New ClassRelFluxoCxItem
                        
            With objRelFluxoCxItem
            
                .dValor = dValor
                .sDescricao = sDescricao
                .iNumParcela = iNumParcela
                .iNumParcelas = iNumParcelas
                
                If iTipo = NATUREZA_TIPO_RECEBIMENTO Then
                    colRec.Add objRelFluxoCxItem
                Else
                    colPag.Add objRelFluxoCxItem
                End If
                
            End With
        
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33333
        
        Loop
        
    End If

    ObterDadosFluxo = SUCESSO

    Exit Function

Erro_ObterDadosFluxo:

    ObterDadosFluxo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Sub Class_Initialize()
    dtUltData = DATA_NULA
End Sub

Private Sub Class_Terminate()
    Call Comando_Fechar(lComando)
End Sub
