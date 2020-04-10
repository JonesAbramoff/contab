VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private colRelMsgAux As New Collection
Private colRelAnotAux As New Collection

Public Function ObterMsg(sRetorno As String, ByVal iTipoDoc As Integer, ByVal lNumIntDoc As Long, ByVal iTipoMsg As Integer, ByVal iTamMax As Integer, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim sMsgAux As String
Dim colMsg As New Collection
Dim bAchou As Boolean
Dim objRelMsgDocAux As ClassRelMsgDocAux
Dim objRelMsgTipoAux As ClassRelMsgTipoAux
Dim objMsg As ClassMensagensRegra, iPosNewLine As Integer, iPosEspaco As Integer, iPosAnt As Integer

On Error GoTo Erro_ObterMsg

    sRetorno = ""

    'Obtem a mensagem (se não estiver em memória tem que ler)
    bAchou = False
    For Each objRelMsgDocAux In colRelMsgAux
        If lNumIntDoc = objRelMsgDocAux.lNumIntDoc And iTipoDoc = objRelMsgDocAux.iTipoDoc Then
            bAchou = True
            Exit For
        End If
    Next
    
    If Not bAchou Then
    
        Set objRelMsgDocAux = New ClassRelMsgDocAux
        
        objRelMsgDocAux.iTipoDoc = iTipoDoc
        objRelMsgDocAux.lNumIntDoc = lNumIntDoc
        
        lErro = CF("MensagensRegra_Le", iTipoDoc, lNumIntDoc, colMsg)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        If colMsg.Count > 0 Then
        
            For Each objMsg In colMsg
                
                If Len(Trim(objMsg.sMensagem)) > 0 Then
                    bAchou = False
                    For Each objRelMsgTipoAux In objRelMsgDocAux.colTipoMsg
                        If objMsg.iTipoMsg = objRelMsgTipoAux.iTipoMsg Then
                            objRelMsgTipoAux.sMsg = objRelMsgTipoAux.sMsg + objMsg.sMensagem
                            objRelMsgTipoAux.sMsgAux = objRelMsgTipoAux.sMsg
                            bAchou = True
                            Exit For
                        End If
                    Next
                    If Not bAchou Then
                        Set objRelMsgTipoAux = New ClassRelMsgTipoAux
                        objRelMsgDocAux.colTipoMsg.Add objRelMsgTipoAux
                        objRelMsgTipoAux.iTipoMsg = objMsg.iTipoMsg
                        objRelMsgTipoAux.sMsg = objMsg.sMensagem
                        objRelMsgTipoAux.sMsgAux = objRelMsgTipoAux.sMsg
                    End If
                End If
            
            Next
            
        End If
        
        colRelMsgAux.Add objRelMsgDocAux
    
    End If
    
    'Com os dados desse tipo de documento já em memória verifica a mensagem para o tipo
    bAchou = False
    For Each objRelMsgTipoAux In objRelMsgDocAux.colTipoMsg
        If iTipoMsg = objRelMsgTipoAux.iTipoMsg Then
            bAchou = True
            Exit For
        End If
    Next
    
    'Se existe alguma mensagem para esse Doc\Tipo Msg
    If bAchou Then
    
        If iLinha > objRelMsgTipoAux.iLinhaUlt Then
        
            If Len(Trim(objRelMsgTipoAux.sMsgAux)) > 0 Then
                
                iPosNewLine = InStr(1, objRelMsgTipoAux.sMsgAux, "|")
                If iTamMax = 0 Then iTamMax = 255
                
                'Se tem tamanho máximo de retorno e é menor que a quebra de linha
                'Tem que retornar até o último espaço antes do tamanho máximo
                If iTamMax < iPosNewLine Or (iPosNewLine = 0 And iTamMax < Len(objRelMsgTipoAux.sMsgAux)) Then
                    
                    'Corta a mensagem até o máximo permitido
                    sMsgAux = left(objRelMsgTipoAux.sMsgAux, iTamMax)
                    
                    'Busca o último espaço antes do corte
                    iPosAnt = 0
                    iPosEspaco = InStr(1, sMsgAux, " ")
                    Do While iPosEspaco <> 0
                        iPosAnt = iPosEspaco
                        iPosEspaco = InStr(iPosAnt + 1, sMsgAux, " ")
                    Loop
                    
                    'Se achou algum espaço Corta a mensagem até ele
                    If iPosAnt <> 0 Then
                        sMsgAux = left(sMsgAux, iPosAnt - 1)
                    End If
                ElseIf iPosNewLine = 0 Then
                'Se a mensagem está inteira e menor que o tamanho máximo
                    sMsgAux = objRelMsgTipoAux.sMsgAux
                Else
                'Se não retorna até a quebra de linha
                    sMsgAux = left(objRelMsgTipoAux.sMsgAux, iPosNewLine - 1)
                End If
                
                objRelMsgTipoAux.iLinhaUlt = objRelMsgTipoAux.iLinhaUlt + 1
                
                If Len(Trim(objRelMsgTipoAux.sMsgAux)) > Len(sMsgAux) + 2 Then
                    objRelMsgTipoAux.sMsgAux = Mid(objRelMsgTipoAux.sMsgAux, Len(sMsgAux) + 2)
                Else
                    objRelMsgTipoAux.sMsgAux = ""
                End If
                
                objRelMsgTipoAux.colMsgLinha.Add sMsgAux
                
                sRetorno = sMsgAux
            
            End If
            
        Else
        
            If objRelMsgTipoAux.colMsgLinha.Count >= iLinha Then
                sRetorno = objRelMsgTipoAux.colMsgLinha.Item(iLinha)
            End If
        
        End If
    
    End If
    
    ObterMsg = SUCESSO
    
    Exit Function
    
Erro_ObterMsg:

    ObterMsg = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153265)

    End Select

    Exit Function
    
End Function

Public Function ObterAnot(sRetorno As String, ByVal iTipoDoc As Integer, ByVal sID As String, ByVal iTamMax As Integer, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim sMsgAux As String, bAchou As Boolean
Dim objAnotacoes As ClassAnotacoes, sTextoAux As String
Dim iPosNewLine As Integer, iPosEspaco As Integer, iPosAnt As Integer
Dim objRelMsgDocAux As ClassRelMsgDocAux
Dim objRelMsgTipoAux As ClassRelMsgTipoAux

On Error GoTo Erro_ObterAnot

    sRetorno = ""

    'Obtem a mensagem (se não estiver em memória tem que ler)
    bAchou = False
    For Each objRelMsgDocAux In colRelAnotAux
        If sID = objRelMsgDocAux.sID And iTipoDoc = objRelMsgDocAux.iTipoDoc Then
            Set objRelMsgTipoAux = objRelMsgDocAux.colTipoMsg(1)
            bAchou = True
            Exit For
        End If
    Next
    
    If Not bAchou Then
    
        Set objAnotacoes = New ClassAnotacoes
        
        objAnotacoes.iTipoDocOrigem = iTipoDoc
        objAnotacoes.sID = sID
        
        lErro = CF("Anotacoes_Le_OrigemID", objAnotacoes)
        If lErro <> SUCESSO And lErro <> 119844 Then gError ERRO_SEM_MENSAGEM
        
        sTextoAux = ""
        lErro = CF("AnotacoesLinha_Le_Texto1", sTextoAux, objAnotacoes.lNumIntDoc)
        If lErro <> SUCESSO And lErro <> 119875 Then gError ERRO_SEM_MENSAGEM
        
        objAnotacoes.sTextoCompleto = sTextoAux
        
        Set objRelMsgDocAux = New ClassRelMsgDocAux
        
        objRelMsgDocAux.sID = sID
        objRelMsgDocAux.iTipoDoc = iTipoDoc
        Set objRelMsgTipoAux = New ClassRelMsgTipoAux
        objRelMsgDocAux.colTipoMsg.Add objRelMsgTipoAux
        objRelMsgTipoAux.sMsg = objAnotacoes.sTextoCompleto
        
        objRelMsgTipoAux.sMsg = Replace(objRelMsgTipoAux.sMsg, vbNewLine, "|")
        objRelMsgTipoAux.sMsg = Replace(objRelMsgTipoAux.sMsg, Chr(10), "|")
        objRelMsgTipoAux.sMsg = Replace(objRelMsgTipoAux.sMsg, Chr(13), "|")
        
        objRelMsgTipoAux.sMsgAux = objRelMsgTipoAux.sMsg
                
        colRelAnotAux.Add objRelMsgDocAux
    
    End If
 
    If iLinha > objRelMsgTipoAux.iLinhaUlt Then
    
        If Len(Trim(objRelMsgTipoAux.sMsgAux)) > 0 Then
            
            iPosNewLine = InStr(1, objRelMsgTipoAux.sMsgAux, "|")
            If iTamMax = 0 Then iTamMax = 255
            
            'Se tem tamanho máximo de retorno e é menor que a quebra de linha
            'Tem que retornar até o último espaço antes do tamanho máximo
            If iTamMax < iPosNewLine Or (iPosNewLine = 0 And iTamMax < Len(objRelMsgTipoAux.sMsgAux)) Then
                
                'Corta a mensagem até o máximo permitido
                sMsgAux = left(objRelMsgTipoAux.sMsgAux, iTamMax)
                
                'Busca o último espaço antes do corte
                iPosAnt = 0
                iPosEspaco = InStr(1, sMsgAux, " ")
                Do While iPosEspaco <> 0
                    iPosAnt = iPosEspaco
                    iPosEspaco = InStr(iPosAnt + 1, sMsgAux, " ")
                Loop
                
                'Se achou algum espaço Corta a mensagem até ele
                If iPosAnt <> 0 Then
                    sMsgAux = left(sMsgAux, iPosAnt - 1)
                End If
            ElseIf iPosNewLine = 0 Then
            'Se a mensagem está inteira e menor que o tamanho máximo
                sMsgAux = objRelMsgTipoAux.sMsgAux
            Else
            'Se não retorna até a quebra de linha
                sMsgAux = left(objRelMsgTipoAux.sMsgAux, iPosNewLine - 1)
            End If
            
            objRelMsgTipoAux.iLinhaUlt = objRelMsgTipoAux.iLinhaUlt + 1
            
            If Len(Trim(objRelMsgTipoAux.sMsgAux)) > Len(sMsgAux) + 2 Then
                objRelMsgTipoAux.sMsgAux = Mid(objRelMsgTipoAux.sMsgAux, Len(sMsgAux) + 2)
            Else
                objRelMsgTipoAux.sMsgAux = ""
            End If
            
            objRelMsgTipoAux.colMsgLinha.Add sMsgAux
            
            sRetorno = sMsgAux
        
        End If
        
    Else
    
        If objRelMsgTipoAux.colMsgLinha.Count >= iLinha Then
            sRetorno = objRelMsgTipoAux.colMsgLinha.Item(iLinha)
        End If
    
    End If
    
    ObterAnot = SUCESSO
    
    Exit Function
    
Erro_ObterAnot:

    ObterAnot = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153265)

    End Select

    Exit Function
    
End Function

Private Sub Class_Terminate()
    Set colRelMsgAux = Nothing
    Set colRelAnotAux = Nothing
End Sub
