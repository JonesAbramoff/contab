VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassUltEmpenho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'para obter a data do ultimo empenho de um produto
'como nao posso garantir a ordem com que os produtos serao passados nao vou poder fazer um SELECT unico

Private lComando As Long

'para evitar que o mesmo produto seja pesquisado seguidas vezes
Private sUltProduto As String
Private dtUltData As Date

Public Function ObterUltEmpenhoProd(sProduto As String, dtDataUltEmpenho As Date) As Long
'para obter a data do ultimo empenho de um produto

Dim lErro As Long
On Error GoTo Erro_ObterUltEmpenhoProd

    If sProduto = sUltProduto Then
    
        dtDataUltEmpenho = dtUltData
    
    Else
    
        lErro = CF("Produto_ObterDataUltEmpenho",sProduto, dtDataUltEmpenho, lComando)
        If lErro <> SUCESSO Then Error 41554
        
        sUltProduto = sProduto
        dtUltData = dtDataUltEmpenho
        
    End If

    ObterUltEmpenhoProd = SUCESSO
    
    Exit Function
    
Erro_ObterUltEmpenhoProd:

    ObterUltEmpenhoProd = Err
    
    Select Case Err

        Case 41554
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154188)

    End Select

    Exit Function
    
End Function

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)
    
End Sub
