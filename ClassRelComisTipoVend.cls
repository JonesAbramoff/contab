VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelComisTipoVend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Le_Comissoes_PorGeracoes(dTotalComissoes As Double, dtDataIni As Date, dtDataFinal As Date, iTipoVendedor As Integer)

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer

On Error GoTo Erro_Le_Comissoes_PorGeracoes
    
    'Abre o lComando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 84655
    
    If dtDataFinal <> DATA_NULA Then
        lErro = Comando_Executar(lComando, "SELECT SUM(Valor) FROM Comissoes, Vendedores  WHERE Vendedores.Tipo=? AND Comissoes.DataGeracao <= ?  AND Comissoes.DataGeracao >= ? AND Comissoes.CodVendedor = Vendedores.Codigo", dTotalComissoes, iTipoVendedor, dtDataFinal, dtDataIni)
    Else
        lErro = Comando_Executar(lComando, "SELECT SUM(Valor) FROM Comissoes, Vendedores  WHERE Vendedores.Tipo=? AND Comissoes.DataGeracao >= ? AND Comissoes.CodVendedor = Vendedores.Codigo", dTotalComissoes, iTipoVendedor, dtDataIni)
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 84657
    
    'Busca o primeiro elemento em Vendedores
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 84658

    'Fecha os lComandos
    Call Comando_Fechar(lComando)
    
    Le_Comissoes_PorGeracoes = SUCESSO

    Exit Function

Erro_Le_Comissoes_PorGeracoes:

    Le_Comissoes_PorGeracoes = gErr
    
    Select Case gErr
        
        Case 84655
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 84657, 84658
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORES", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153063)
    
    End Select
    
    'Fecha os lComandos
    Call Comando_Fechar(lComando)
  
    
    Exit Function

End Function

Public Function Le_Comissoes_PorGeracoes_Vendedor(dTotalComissoes As Double, dtDataIni As Date, dtDataFinal As Date, iTipoVendedor As Integer, iCodVendedor As Integer) As Long

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer

On Error GoTo Erro_Le_Comissoes_PorGeracoes_Vendedor
    
    'Abre o lComando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 133006
    
    If dtDataFinal <> DATA_NULA Then
        lErro = Comando_Executar(lComando, "SELECT SUM(Valor) FROM Comissoes, Vendedores  WHERE Vendedores.Tipo=? AND Comissoes.DataGeracao <= ?  AND Comissoes.DataGeracao >= ? AND Comissoes.CodVendedor = Vendedores.Codigo AND Comissoes.CodVendedor = ?", dTotalComissoes, iTipoVendedor, dtDataFinal, dtDataIni, iCodVendedor)
    Else
        lErro = Comando_Executar(lComando, "SELECT SUM(Valor) FROM Comissoes, Vendedores  WHERE Vendedores.Tipo=? AND Comissoes.DataGeracao >= ? AND Comissoes.CodVendedor = Vendedores.Codigo AND Comissoes.CodVendedor = ?", dTotalComissoes, iTipoVendedor, dtDataIni, iCodVendedor)
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 133007
    
    'Busca o primeiro elemento em Vendedores
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133008

    'Fecha os lComandos
    Call Comando_Fechar(lComando)
    
    Le_Comissoes_PorGeracoes_Vendedor = SUCESSO

    Exit Function

Erro_Le_Comissoes_PorGeracoes_Vendedor:

    Le_Comissoes_PorGeracoes_Vendedor = gErr
    
    Select Case gErr
        
        Case 133006
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 133007, 133008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORES", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153064)
    
    End Select
    
    'Fecha os lComandos
    Call Comando_Fechar(lComando)
  
    
    Exit Function

End Function

