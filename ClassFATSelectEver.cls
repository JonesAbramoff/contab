VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function AlmoxarifadosFilial_Le_Quantidade(ByVal iFilialEmpresa As Integer, lNumAlmoxarifados As Long) As Long
'Retorna o número de Almoxarifados da FilialEmpresa

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_AlmoxarifadosFilial_Le_Quantidade

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 39366
    
    If iFilialEmpresa = 52 Then iFilialEmpresa = 51
    
    If iFilialEmpresa = 2 Then iFilialEmpresa = 1
    
    'Busca o número de almoxarifados em Almoxarifado
    lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM Almoxarifado WHERE FilialEmpresa = ? ", lNumAlmoxarifados, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 39367
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 39368
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    AlmoxarifadosFilial_Le_Quantidade = SUCESSO
    
    Exit Function
    
Erro_AlmoxarifadosFilial_Le_Quantidade:
    
    AlmoxarifadosFilial_Le_Quantidade = Err
    
    Select Case Err
        
        Case 39366
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err, Error$)
    
        Case 39367, 39368
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADOS", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 179020)
            
    End Select
    
    Call Comando_Fechar(lComando)
        
    Exit Function
        
End Function

Sub TabelaPrecoItem_Valor_GotFocus(objValor As Object)

    objValor.SelStart = 0
    objValor.SelLength = Len(objValor.Text)

End Sub

Function Serie_FilialEmpresa_Customiza(objSerie As ClassSerie) As Long

    If objSerie.sSerie <> "1" Then objSerie.iFilialEmpresa = FilialEmpresa_ConvFRFO(objSerie.iFilialEmpresa)

    Serie_FilialEmpresa_Customiza = SUCESSO

End Function

Function NFiscal_FilialEmpresa_Customiza(objNFiscal As ClassNFiscal) As Long

    If objNFiscal.sSerie <> "1" Then objNFiscal.iFilialEmpresa = FilialEmpresa_ConvFRFO(objNFiscal.iFilialEmpresa)
    
    NFiscal_FilialEmpresa_Customiza = SUCESSO

End Function

Function LiberacaoBloqueio_FilialEmpresa(objLiberacaoBloqueios As ClassLiberacaoBloqueios) As Long

    objLiberacaoBloqueios.iFilialEmpresa = EMPRESA_TODA
    
    LiberacaoBloqueio_FilialEmpresa = SUCESSO

End Function

Function Vendedor_Comissao_Customiza(objVendedor As ClassVendedor, dPercComissao As Double) As Long
    
    dPercComissao = objVendedor.dPercComissao
    
    Vendedor_Comissao_Customiza = SUCESSO

End Function

Function Vendedor_Comissao_Customiza1(objVendedor As ClassVendedor, objComissaoNF As ClassComissaoNF, objFilialCliente As ClassFilialCliente) As Long
    
    objComissaoNF.dPercentual = objVendedor.dPercComissao
    
    Vendedor_Comissao_Customiza1 = SUCESSO

End Function

'####################################################
'Inserido por Wagner 19/05/2006
Function FilialEmpresa_Le_Filtro(sFiltro As String) As Long
    
    sFiltro = " AND FilialEmpresa > " & DELTA_FILIALREAL_OFICIAL & " "
    
    FilialEmpresa_Le_Filtro = SUCESSO
        
End Function
'####################################################

'####################################################
'Inserido por Wagner 22/05/2006
Function FilialEmpresa_Le_Filtro_Filial(sFiltro As String) As Long

    If giFilialEmpresa > DELTA_FILIALREAL_OFICIAL Then
    
        sFiltro = " AND FilialEmpresa > " & DELTA_FILIALREAL_OFICIAL & " "
        
    Else
    
        sFiltro = ""
        
    End If
    
    FilialEmpresa_Le_Filtro_Filial = SUCESSO
        
End Function
'####################################################

Function Cod_Nomes_Le_FilEmp(ByVal colCodigoNome As AdmColCodigoNome) As Long

Dim lErro As Long
Dim objFiliais As AdmFiliais

On Error GoTo Erro_Cod_Nomes_Le_FilEmp

    For Each objFiliais In gcolFiliais
        
        If objFiliais.iInativa = 0 And objFiliais.iCodFilial <> EMPRESA_TODA Then
        
            If (giFilialEmpresa > DELTA_FILIALREAL_OFICIAL And objFiliais.iCodFilial > DELTA_FILIALREAL_OFICIAL) Or _
                (giFilialEmpresa < DELTA_FILIALREAL_OFICIAL And objFiliais.iCodFilial < DELTA_FILIALREAL_OFICIAL) Then
                Call colCodigoNome.Add(objFiliais.iCodFilial, objFiliais.sNome)
            End If
        
        End If
        
    Next
    
    Cod_Nomes_Le_FilEmp = SUCESSO
     
    Exit Function
    
Erro_Cod_Nomes_Le_FilEmp:

    Cod_Nomes_Le_FilEmp = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130531)
     
    End Select
     
    Exit Function

End Function

