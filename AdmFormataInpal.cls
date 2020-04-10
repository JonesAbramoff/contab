VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmFormataInpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Produto_Formata(ByVal sProduto As String, sProdutoFormatado As String, iProdutoPreenchido As Integer, Optional ByVal bSemMsgErro As Boolean = False) As Long
'Devolve o produto no formato do BD: zeros a esquerda em segmentos numericos e brancos a direita para segmentos alfabeticos
'A String de Entrada vem com os caracteres delimitadores e brancos(se for o caso). O comprimento dela é o comprimento total da máscara.
'iProdutoPreenchido retornará PRODUTO_VAZIO se nenhum segmento do produto for preenchido,
'caso contrário retornará PRODUTO_PREENCHIDO
    
Dim lErro As Long
Dim iNumSegmentos As Integer
Dim iInd1 As Integer
Dim iInd2 As Integer
Dim objSegmento As New ClassSegmento
Dim sProdutoBruto As String
Dim sSegmentoBruto As String 'Segmento sendo formatado
Dim sSegmentoFormatado As String 'Segmento sendo formatado
Dim sCaracter As String 'caracter sendo tratado
Dim iRestoVazio As Integer
Dim iSegmentoVazio As Integer

On Error GoTo Erro_Produto_Formata
            
    iRestoVazio = SEGMENTOS_RESTANTES_CHEIOS
    iProdutoPreenchido = PRODUTO_VAZIO
    sProdutoBruto = sProduto
    sProdutoFormatado = ""
    iNumSegmentos = NumSegProduto()

    For iInd1 = 1 To iNumSegmentos
        
        lErro = SegmentoProduto(iInd1, objSegmento)
        If lErro <> SUCESSO Then Error 6072
        
        sSegmentoBruto = left(sProdutoBruto, objSegmento.iTamanho)
        sProdutoBruto = Mid(sProdutoBruto, objSegmento.iTamanho + Len(objSegmento.sDelimitador) + 1)
        sSegmentoFormatado = ""

        iSegmentoVazio = SEGMENTO_VAZIO

        Select Case objSegmento.iTipo
        
            Case SEGMENTO_NUMERICO

                For iInd2 = 1 To objSegmento.iTamanho
    
                    sCaracter = Mid(sSegmentoBruto, iInd2, 1)
    
                    If sCaracter = " " Or sCaracter = "" Then
                        sSegmentoFormatado = "0" & sSegmentoFormatado
                        
                    ElseIf sCaracter = "0" Then
                        sSegmentoFormatado = sSegmentoFormatado & sCaracter
                        
                    ElseIf sCaracter = "-" Then
                        Error 55692
                    Else
                        iSegmentoVazio = SEGMENTO_CHEIO
                        sSegmentoFormatado = sSegmentoFormatado & sCaracter
                    End If
    
                Next

            Case SEGMENTO_ALFANUMERICO, SEGMENTO_ASCII

                For iInd2 = objSegmento.iTamanho To 1 Step -1
    
                    sCaracter = Mid(sSegmentoBruto, iInd2, 1)
    
                    If sCaracter = " " Or sCaracter = "" Then
                        sSegmentoFormatado = sSegmentoFormatado & " "
                    Else
                        iSegmentoVazio = SEGMENTO_CHEIO
                        sSegmentoFormatado = UCase(sCaracter) & sSegmentoFormatado
                    End If
    
                Next

        End Select

        If iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS And iSegmentoVazio <> SEGMENTO_VAZIO Then Error 20736
            
        If iSegmentoVazio = SEGMENTO_VAZIO Then
            iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS
        Else
            iProdutoPreenchido = PRODUTO_PREENCHIDO
        End If


        sProdutoFormatado = sProdutoFormatado & sSegmentoFormatado

    Next
    
    If iProdutoPreenchido = PRODUTO_VAZIO Then sProdutoFormatado = ""
    
    Produto_Formata = SUCESSO
    
    Exit Function
    
Erro_Produto_Formata:

    Produto_Formata = Err

    Select Case Err

        Case 20736
            If Not bSemMsgErro Then Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEG_MEIO_NAO_PREENCHIDOS", Err)

        Case 25035 'Vazio porque o erro já foi tratado

        Case 55692
            If Not bSemMsgErro Then Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEG_NUM_CARACTER_INVALIDO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select

    Exit Function
    
End Function

