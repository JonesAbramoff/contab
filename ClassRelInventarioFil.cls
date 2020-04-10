VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelInventarioFil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim lComando As Long
Dim sProdutoAnt As String
Dim iFilialEmpresaAnt As Integer
Dim dValorAnt As Double

'para calculo de estoque numa data
Private sProdutoAntInv As String
Private dtDataAntInv As Date

Private dQuantidadeNossaInv As Double
Private dQuantidadeEmTerc As Double
Private dQuantidadeDeTerc As Double

Public Function ObterSaldoProdDtFil(dQuantidade As Double, sProduto As String, dtData As Date, iFilialEmpresa As Integer) As Long
'Para obter o saldo em estoque da propria empresa de um produto em uma data

Dim lErro As Long, objRegInventarioFil As New ClassRegInventario

On Error GoTo Erro_ObterSaldoProdDtFil

    If (sProdutoAntInv <> sProduto Or dtDataAntInv <> dtData) Then
    
        With objRegInventarioFil
            .iFilialEmpresa = iFilialEmpresa
            .sProduto = sProduto
            .iAlmoxarifado = 0
            .dtData = dtData
        End With
        
        lErro = CF("RegInventario_Obter_Quantidades",objRegInventarioFil)
        If lErro <> SUCESSO Then gError 90371
        
        sProdutoAntInv = sProduto
        dtDataAntInv = dtData
        
        objRegInventarioFil.dQtdeDeTercConosco = objRegInventarioFil.dQuantBenef3 + objRegInventarioFil.dQuantConserto3 + objRegInventarioFil.dQuantConsig3 + objRegInventarioFil.dQuantDemo3 + objRegInventarioFil.dQuantOutras3
        objRegInventarioFil.dQtdeNossaEmTerc = objRegInventarioFil.dQuantBenef + objRegInventarioFil.dQuantConserto + objRegInventarioFil.dQuantConsig + objRegInventarioFil.dQuantDemo + objRegInventarioFil.dQuantOutras
        
        dQuantidadeNossaInv = objRegInventarioFil.dQuantidadeUMEstoque
        dQuantidadeEmTerc = objRegInventarioFil.dQtdeNossaEmTerc
        dQuantidadeDeTerc = objRegInventarioFil.dQtdeDeTercConosco
        
    End If
        
    dQuantidade = dQuantidadeNossaInv - dQuantidadeDeTerc
        
    ObterSaldoProdDtFil = SUCESSO
     
    Exit Function
    
Erro_ObterSaldoProdDtFil:

    ObterSaldoProdDtFil = gErr
     
    Select Case gErr
          
        Case 90371
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153171)
     
    End Select
     
    Exit Function

End Function

Public Function ObterSaldoProdDtFilDeTerc(dQuantidade As Double, sProduto As String, dtData As Date, iFilialEmpresa As Integer) As Long
'Para obter o saldo de estoque de 3os em um produto-almoxarifado em uma data

Dim lErro As Long, objRegInventarioFil As New ClassRegInventario

On Error GoTo Erro_ObterSaldoProdDtFilDeTerc

    If (sProdutoAntInv <> sProduto Or dtDataAntInv <> dtData) Then
    
        With objRegInventarioFil
            .iFilialEmpresa = iFilialEmpresa
            .sProduto = sProduto
            .iAlmoxarifado = 0
            .dtData = dtData
        End With
        
        lErro = CF("RegInventario_Obter_Quantidades",objRegInventarioFil)
        If lErro <> SUCESSO Then gError 90372
        
        sProdutoAntInv = sProduto
        dtDataAntInv = dtData
        
        objRegInventarioFil.dQtdeDeTercConosco = objRegInventarioFil.dQuantBenef3 + objRegInventarioFil.dQuantConserto3 + objRegInventarioFil.dQuantConsig3 + objRegInventarioFil.dQuantDemo3 + objRegInventarioFil.dQuantOutras3
        objRegInventarioFil.dQtdeNossaEmTerc = objRegInventarioFil.dQuantBenef + objRegInventarioFil.dQuantConserto + objRegInventarioFil.dQuantConsig + objRegInventarioFil.dQuantDemo + objRegInventarioFil.dQuantOutras
        
        dQuantidadeNossaInv = objRegInventarioFil.dQuantidadeUMEstoque
        dQuantidadeEmTerc = objRegInventarioFil.dQtdeNossaEmTerc
        dQuantidadeDeTerc = objRegInventarioFil.dQtdeDeTercConosco
            
    End If
    
    dQuantidade = dQuantidadeDeTerc
        
    ObterSaldoProdDtFilDeTerc = SUCESSO
     
    Exit Function
    
Erro_ObterSaldoProdDtFilDeTerc:

    ObterSaldoProdDtFilDeTerc = gErr
     
    Select Case gErr
          
        Case 90372
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153172)
     
    End Select
     
    Exit Function

End Function

Public Function ObterSaldoProdDtFilEmTerc(dQuantidade As Double, sProduto As String, dtData As Date, iFilialEmpresa As Integer) As Long
'Para obter o saldo de estoque em 3os alocada a um produto-almoxarifado em uma data

Dim lErro As Long, objRegInventarioFil As New ClassRegInventario

On Error GoTo Erro_ObterSaldoProdDtFilEmTerc

    If (sProdutoAntInv <> sProduto Or dtDataAntInv <> dtData) Then
    
        With objRegInventarioFil
            .iFilialEmpresa = iFilialEmpresa
            .sProduto = sProduto
            .iAlmoxarifado = 0
            .dtData = dtData
        End With
        
        lErro = CF("RegInventario_Obter_Quantidades",objRegInventarioFil)
        If lErro <> SUCESSO Then gError 90373
        
        sProdutoAntInv = sProduto
        dtDataAntInv = dtData
        
        objRegInventarioFil.dQtdeDeTercConosco = objRegInventarioFil.dQuantBenef3 + objRegInventarioFil.dQuantConserto3 + objRegInventarioFil.dQuantConsig3 + objRegInventarioFil.dQuantDemo3 + objRegInventarioFil.dQuantOutras3
        objRegInventarioFil.dQtdeNossaEmTerc = objRegInventarioFil.dQuantBenef + objRegInventarioFil.dQuantConserto + objRegInventarioFil.dQuantConsig + objRegInventarioFil.dQuantDemo + objRegInventarioFil.dQuantOutras
        
        dQuantidadeNossaInv = objRegInventarioFil.dQuantidadeUMEstoque
        dQuantidadeEmTerc = objRegInventarioFil.dQtdeNossaEmTerc
        dQuantidadeDeTerc = objRegInventarioFil.dQtdeDeTercConosco
        
    End If
        
    dQuantidade = dQuantidadeEmTerc
        
    ObterSaldoProdDtFilEmTerc = SUCESSO
     
    Exit Function
    
Erro_ObterSaldoProdDtFilEmTerc:

    ObterSaldoProdDtFilEmTerc = gErr
     
    Select Case gErr
          
        Case 90373
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153173)
     
    End Select
     
    Exit Function

End Function

