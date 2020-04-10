VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelInfoAdicDocItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private giUltTipodoc As Integer
Private glUltNumIntDocItem As Long
Private gobjInfoAdicDocItem As ClassInfoAdicDocItem
Private gcolMsgAux As Collection

Private Sub QuebraTextoEmLinhas(ByVal sMsg As String, ByVal iTamMax As Integer, ByVal colLinhas As Collection)

Dim sMsgAux As String, iPOS As Integer, iPosNewLine As Integer
Dim iPosEspaco As Integer, iPosAnt As Integer

    sMsg = Trim(sMsg)
    
    sMsg = Replace(sMsg, vbNewLine, "|")
    sMsg = Replace(sMsg, Chr(10), "|")
    sMsg = Replace(sMsg, Chr(13), "|")
    
    sMsgAux = sMsg
    
    Do While Len(sMsgAux) <> 0
    
        iPosNewLine = InStr(1, sMsg, "|")
        If iTamMax = 0 Then iTamMax = 255
        
        'Se tem tamanho máximo de retorno e é menor que a quebra de linha
        'Tem que retornar até o último espaço antes do tamanho máximo
        If iTamMax < iPosNewLine Or (iPosNewLine = 0 And iTamMax < Len(sMsg)) Then
            
            'Corta a mensagem até o máximo permitido
            sMsgAux = left(sMsg, iTamMax)
                
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
                sMsg = Mid(sMsg, Len(sMsgAux) + 2)
            Else
                sMsg = Mid(sMsg, Len(sMsgAux) + 1)
            End If
            
        ElseIf iPosNewLine = 0 Then
        'Se a mensagem está inteira e menor que o tamanho máximo
            sMsgAux = sMsg
            sMsg = ""
        Else
        'Se não retorna até a quebra de linha
            sMsgAux = left(sMsg, iPosNewLine - 1)
            sMsg = Mid(sMsg, Len(sMsgAux) + 2)
        End If
        
        colLinhas.Add RTrim(sMsgAux)
    
        sMsgAux = sMsg
    
    Loop

End Sub

Public Function ObterMsg(sRetorno As String, ByVal iTipoDoc As Integer, ByVal lNumIntDocItem As Long, ByVal iTamMax As Integer, ByVal iLinha As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_ObterMsg

    sRetorno = ""
    
    If lNumIntDocItem <> 0 And iLinha > 0 Then
    
        If iTipoDoc <> giUltTipodoc Or glUltNumIntDocItem <> lNumIntDocItem Then
        
            Set gobjInfoAdicDocItem = New ClassInfoAdicDocItem
        
            gobjInfoAdicDocItem.iTipoDoc = iTipoDoc
            gobjInfoAdicDocItem.lNumIntDocItem = lNumIntDocItem
            
            lErro = CF("InfoAdicDocItem_Le", gobjInfoAdicDocItem)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                                
            Set gcolMsgAux = New Collection
            
            Call QuebraTextoEmLinhas(gobjInfoAdicDocItem.sMsg, iTamMax, gcolMsgAux)
            
            giUltTipodoc = iTipoDoc
            glUltNumIntDocItem = lNumIntDocItem
            
        End If
        
        If iLinha <= gcolMsgAux.Count Then sRetorno = gcolMsgAux.Item(iLinha)

    End If

    ObterMsg = SUCESSO
    
    Exit Function
    
Erro_ObterMsg:

    ObterMsg = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201455)

    End Select
    
    Exit Function

End Function


