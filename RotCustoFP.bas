Attribute VB_Name = "RotCustoFP"
Option Explicit

Sub CustoDiretoFabricacao_Calcula2(dAcumQuantMedia As Double, tPrev As typePrevVendaMensal2, ByVal iMesIni As Integer, ByVal iMesFim As Integer)
'acumular qtde de meses e qtde do produto para os meses validos
    '(obs.: os meses "validos" são aqueles a partir do 1o mes em que a dataatualizacaomes for <> DATA_NULA e estiver dentro do periodo)
Dim iMes As Integer, bAchou As Boolean, iAcumMeses As Integer, dAcumQuant As Double

    bAchou = False

    With tPrev
    
        For iMes = 1 To 12
        
            If bAchou = False Then
            
                If .adtDataAtualizacao(iMes) <> DATA_NULA Then
                
                    bAchou = True
                    
                    If iMes >= iMesIni And iMes <= iMesFim Then
                    
                        iAcumMeses = iAcumMeses + 1
                        dAcumQuant = dAcumQuant + .adQuantidade(iMes)
                    
                    End If
                    
                End If
                
            Else
            
                If iMes >= iMesIni And iMes <= iMesFim Then
                
                    iAcumMeses = iAcumMeses + 1
                    dAcumQuant = dAcumQuant + .adQuantidade(iMes)
                
                End If
                
            End If
        
        Next
        
    End With
    
    If iAcumMeses <> 0 Then
    
        dAcumQuantMedia = ArredondaMod(dAcumQuantMedia + dAcumQuant / iAcumMeses, 0)
        
    End If
    
End Sub

Function ArredondaMod(dVal As Double, ByVal iCasas As Integer) As Double

    ArredondaMod = Round(dVal + 0.00001, iCasas)
    
End Function

