VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelClassiFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private ColClassFiscal As Object
Private ColLetras As Object
Private sUltimaLetra As String
Private lNumIntDocAnt As Long
Private lComando As Long

Private Sub InicializaColecoes()

    Set ColClassFiscal = CreateObject("Scripting.Dictionary")
    Set ColLetras = CreateObject("Scripting.Dictionary")
    sUltimaLetra = ""
    
End Sub

Private Sub Class_Initialize()
    Call InicializaColecoes
End Sub

Public Sub Obter_Classificacao_Fiscal(sClass As String, sLetra As String)
    
    'Se existir na colecao
    If ColClassFiscal.Exists(sLetra) = True Then
        sClass = Mid(ColClassFiscal.Item(sLetra), 1, 4) & "." & Mid(ColClassFiscal.Item(sLetra), 5, 2) & "." & Mid(ColClassFiscal.Item(sLetra), 7, 4)
    Else
        sClass = ""
    End If
    
End Sub

Public Sub Obter_Letra(sLetra As String, sClass As String, sLetraVariavel As String, ByVal lNumIntDoc As Long)
'Passa a Classificacao e a Primeira Letra que eu vou utilizar
    
    If lNumIntDocAnt <> lNumIntDoc Then
        Call InicializaColecoes
        lNumIntDocAnt = lNumIntDoc
    End If
    
    'Se não existir na colecao
    If ColLetras.Exists(sClass) = False Then
        
        'Se é a Primeira Letra
        If sUltimaLetra = "" Then
            'Repassa a Letra e quarda a última
            sUltimaLetra = sLetraVariavel
            sLetra = sLetraVariavel
        Else
            sUltimaLetra = Chr(Asc(Trim(sUltimaLetra)) + 1)
            sLetra = sUltimaLetra
        End If
        
        ColClassFiscal.Add sUltimaLetra, sClass
        ColLetras.Add sClass, sUltimaLetra
        
    Else
    
        sLetra = ColLetras.Item(sClass)
    
    End If
    'MsgBox sLetra
    
End Sub

Public Function Obter_Letras_NF(ByVal lNumIntDoc As Long) As Long

Dim lErro As Long
Dim objProduto As ClassProduto
Dim colProdutos As New Collection
Dim sLetra As String, sClass As String, sLetraVar As String

On Error GoTo Erro_Obter_Letras_NF

    'MsgBox CStr(lNumIntDoc)
    If lNumIntDocAnt <> lNumIntDoc Then
        
        lErro = CF("NF_Obter_Letras", lComando, lNumIntDoc, colProdutos)
        If lErro <> SUCESSO Then gError 60458
        
        For Each objProduto In colProdutos
            'MsgBox sClass
            sClass = objProduto.sIPICodigo
            sLetraVar = "A"
            Call Obter_Letra(sLetra, sClass, sLetraVar, lNumIntDoc)
        Next
    
    End If
    
    
    Obter_Letras_NF = SUCESSO
    
    Exit Function
    
Erro_Obter_Letras_NF:
        
    Obter_Letras_NF = gErr
    
    Select Case gErr
        
        Case 60458 'Tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153053)
    
    End Select

    Exit Function
    
End Function

Private Sub Class_Terminate()

    Set ColClassFiscal = Nothing
    Set ColLetras = Nothing
    If lComando <> 0 Then Call Comando_Fechar(lComando)

End Sub
