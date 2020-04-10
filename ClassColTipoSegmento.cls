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
'recebe a descrição do tipo do segmento e retorna o tipo do segmento
'retorna 0 se não achou um tipo de segmento para a descrição fornecida

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
'recebe o tipo do segmento e retorna a descrição
'retorna string vazia se não achou uma descrição para o código fornecido

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
'faz a carga da coleção com os tipos de segmento

Dim objTipoSegmento As ClassTipoSegmento

    Set colTipoSegmento = New Collection
    Set objTipoSegmento = New ClassTipoSegmento
        
    objTipoSegmento.iTipo = SEGMENTO_NUMERICO
    objTipoSegmento.sDescricao = "Numérico"
        
    'armazena em memória
    colTipoSegmento.Add objTipoSegmento
        
    Set objTipoSegmento = New ClassTipoSegmento
        
    objTipoSegmento.iTipo = SEGMENTO_ALFANUMERICO
    objTipoSegmento.sDescricao = "Alfanumérico"
        
    'armazena em memória
    colTipoSegmento.Add objTipoSegmento
        
    Set objTipoSegmento = New ClassTipoSegmento
        
    objTipoSegmento.iTipo = SEGMENTO_ASCII
    objTipoSegmento.sDescricao = "Ascii"
        
    'armazena em memória
    colTipoSegmento.Add objTipoSegmento
        
    
End Sub


Private Sub Class_Terminate()
    'destroi a colecao quando essa class termina
    Set colTipoSegmento = Nothing
End Sub




