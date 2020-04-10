VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassColCodigoSegmento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private colCodigoSegmento As Collection

Public Property Get Item(vntIndexKey As Integer) As ClassCodigoSegmento
  Set Item = colCodigoSegmento(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = colCodigoSegmento.Count
End Property

Public Function Codigo(sDescricao As String) As String
'recebe a descrição do segmento e retorna o codigo alfanumerico que é armazenado no BD.

Dim objCodigoSegmento As ClassCodigoSegmento
Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objCodigoSegmento In colCodigoSegmento
    
        If objCodigoSegmento.sDescricao = sDescricao Then
            Codigo = objCodigoSegmento.sCodigo
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        Codigo = ""
    End If
    
End Function

Public Function Descricao(sCodigo As String) As String
'recebe o codigo alfanumerico do segmento tal qual está armazenado no banco de dados e retorna a descrição
'retorna string vazia se não achou uma descrição para o código fornecido

Dim objCodigoSegmento As ClassCodigoSegmento
Dim iAchou As Integer
    
    iAchou = 0
    
    For Each objCodigoSegmento In colCodigoSegmento
    
        If objCodigoSegmento.sCodigo = sCodigo Then
            Descricao = objCodigoSegmento.sDescricao
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        Descricao = ""
    End If
    
End Function

Private Sub Class_Initialize()
'faz a carga da coleção com os segmentos existentes

Dim objCodigoSegmento As ClassCodigoSegmento

    Set colCodigoSegmento = New Collection
    Set objCodigoSegmento = New ClassCodigoSegmento
        
    objCodigoSegmento.sCodigo = SEGMENTO_CONTA
    objCodigoSegmento.sDescricao = "Conta"
        
    'armazena em memória
    colCodigoSegmento.Add objCodigoSegmento
        
    Set objCodigoSegmento = New ClassCodigoSegmento
    
    objCodigoSegmento.sCodigo = SEGMENTO_CCL
    objCodigoSegmento.sDescricao = "Centro de Custo/Lucro"
        
    'armazena em memória
    colCodigoSegmento.Add objCodigoSegmento
    
    Set objCodigoSegmento = New ClassCodigoSegmento
    
    objCodigoSegmento.sCodigo = SEGMENTO_PRODUTO
    objCodigoSegmento.sDescricao = "Produto"
        
    'armazena em memória
    colCodigoSegmento.Add objCodigoSegmento
    
    '###########################################
    'Inserido por Wagner
    Set objCodigoSegmento = New ClassCodigoSegmento
    
    objCodigoSegmento.sCodigo = SEGMENTO_NATMOVCTA
    objCodigoSegmento.sDescricao = "Natureza"
        
    'armazena em memória
    colCodigoSegmento.Add objCodigoSegmento
    '###########################################
    
    '###########################################
    'Inserido por Wagner
    Set objCodigoSegmento = New ClassCodigoSegmento
    
    objCodigoSegmento.sCodigo = SEGMENTO_PROJETO
    objCodigoSegmento.sDescricao = "Projeto"
        
    'armazena em memória
    colCodigoSegmento.Add objCodigoSegmento
    
    Set objCodigoSegmento = New ClassCodigoSegmento
    
    objCodigoSegmento.sCodigo = SEGMENTO_REFETAPA
    objCodigoSegmento.sDescricao = "Referência"
        
    'armazena em memória
    colCodigoSegmento.Add objCodigoSegmento
    '###########################################
    
End Sub


Private Sub Class_Terminate()
    'destroi a colecao quando essa class termina
    Set colCodigoSegmento = Nothing
End Sub




