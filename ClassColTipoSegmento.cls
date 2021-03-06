VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassColTipoSegmento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private colTipoSegmento As Collection


Public Property Get Item(vntIndexKey As Integer) As ClassTipoSegmento
  Set Item = colTipoSegmento(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = colTipoSegmento.Count
End Property

Public Function TipoSegmento(sDescricao As String) As Long
'recebe a descri��o do tipo do segmento e retorna o tipo do segmento
'retorna 0 se n�o achou um tipo de segmento para a descri��o fornecida

Dim objTipoSegmento As ClassTipoSegmento
Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objTipoSegmento In colTipoSegmento
    
        If objTipoSegmento.sDescricao = sDescricao Then
            TipoSegmento = objTipoSegmento.iTipo
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        TipoSegmento = 0
    End If
    
End Function

Public Function Descricao(iTipo As Integer) As String
'recebe o tipo do segmento e retorna a descri��o
'retorna string vazia se n�o achou uma descri��o para o c�digo fornecido

Dim objTipoSegmento As ClassTipoSegmento
Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objTipoSegmento In colTipoSegmento
    
        If objTipoSegmento.iTipo = iTipo Then
            Descricao = objTipoSegmento.sDescricao
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        Descricao = ""
    End If
    
End Function

Private Sub Class_Initialize()
'faz a carga da cole��o com os tipos de segmento

Dim objTipoSegmento As ClassTipoSegmento

    Set colTipoSegmento = New Collection
    Set objTipoSegmento = New ClassTipoSegmento
        
    objTipoSegmento.iTipo = SEGMENTO_NUMERICO
    objTipoSegmento.sDescricao = "Num�rico"
        
    'armazena em mem�ria
    colTipoSegmento.Add objTipoSegmento
        
    Set objTipoSegmento = New ClassTipoSegmento
        
    objTipoSegmento.iTipo = SEGMENTO_ALFANUMERICO
    objTipoSegmento.sDescricao = "Alfanum�rico"
        
    'armazena em mem�ria
    colTipoSegmento.Add objTipoSegmento
        
    Set objTipoSegmento = New ClassTipoSegmento
        
    objTipoSegmento.iTipo = SEGMENTO_ASCII
    objTipoSegmento.sDescricao = "Ascii"
        
    'armazena em mem�ria
    colTipoSegmento.Add objTipoSegmento
        
    
End Sub


Private Sub Class_Terminate()
    'destroi a colecao quando essa class termina
    Set colTipoSegmento = Nothing
End Sub




