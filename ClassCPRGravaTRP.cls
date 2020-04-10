VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCPRGravaTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function RegiaoVenda_Grava_Customizacao(ByVal objRegiaoVenda As ClassRegiaoVenda) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim objRegiaoVendaTRP As ClassRegiaoVendaTRP
Dim iIndice As Integer

On Error GoTo Erro_RegiaoVenda_Grava_Customizacao

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197179
    Next

    If Not (objRegiaoVenda.objInfoUsu Is Nothing) Then
    
        Set objRegiaoVendaTRP = objRegiaoVenda.objInfoUsu
    
        lErro = Comando_Executar(alComando(1), "DELETE FROM TRPRegiaoVenda WHERE Codigo = ?", objRegiaoVenda.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 197180
    
        lErro = Comando_Executar(alComando(2), "INSERT INTO TRPRegiaoVenda (Codigo, CodRegiaoPai) VALUES (?,?)", objRegiaoVenda.iCodigo, objRegiaoVendaTRP.iCodRegiaoPai)
        If lErro <> AD_SQL_SUCESSO Then gError 197181
    
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RegiaoVenda_Grava_Customizacao = SUCESSO

    Exit Function

Erro_RegiaoVenda_Grava_Customizacao:

    RegiaoVenda_Grava_Customizacao = gErr

    Select Case gErr

        Case 197179
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197180
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPREGIAOVENDA", gErr)

        Case 197181
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_TRPREGIAOVENDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197182)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RegiaoVenda_Exclui_Customizado(ByVal objRegiaoVenda As ClassRegiaoVenda) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim objRegiaoVendaTRP As ClassRegiaoVendaTRP
Dim iIndice As Integer

On Error GoTo Erro_RegiaoVenda_Exclui_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197184
    Next

    If Not (objRegiaoVenda.objInfoUsu Is Nothing) Then
    
        Set objRegiaoVendaTRP = objRegiaoVenda.objInfoUsu
    
        lErro = Comando_Executar(alComando(1), "DELETE FROM TRPRegiaoVenda WHERE Codigo = ?", objRegiaoVenda.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 197185
    
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RegiaoVenda_Exclui_Customizado = SUCESSO

    Exit Function

Erro_RegiaoVenda_Exclui_Customizado:

    RegiaoVenda_Exclui_Customizado = gErr

    Select Case gErr

        Case 197184
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197185
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRPREGIAOVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197186)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function FaturaPagar_Grava_BD_Customizado(ByVal objTituloPagar As ClassTituloPagar, ByVal colParcelaPagar As colParcelaPagar, ByVal colNFPag As ColNFsPag) As Long

Dim lErro As Long
Dim objTitulos As ClassTRPTitulos
Dim sTipoDoc As String
Dim sTiposDoc As String
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim objNfsPag As ClassNFsPag
Dim sTextoAux As String
Dim iPos As Integer
Dim colTiposDoc As New Collection
Dim colCampos As New Collection
Dim colSaida As New Collection
Dim sTipoDocAnt As String
Dim lNumIntDocAux As Long

On Error GoTo Erro_FaturaPagar_Grava_BD_Customizado
   
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200237
    Next

    'ALTERAR PARA SÓ GRAVAR SE NÃO TIVER
    lErro = Comando_Executar(alComando(3), "SELECT NumIntDoc FROM TRPTitulos WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ?", lNumIntDocAux, TRP_TIPO_DOC_DESTINO_TITPAG, objTituloPagar.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 200238

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200239

    If lErro <> AD_SQL_SUCESSO Then
            
        For Each objNfsPag In colNFPag
       
            sTipoDoc = String(STRING_TRPTITULOS_TIPODOC, 0)
       
            lErro = Comando_Executar(alComando(1), "SELECT TipoDoc FROM TRPTitulos WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ?", sTipoDoc, TRP_TIPO_DOC_DESTINO_NFSPAG, objNfsPag.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 200238
        
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200239
            
            If lErro = AD_SQL_SUCESSO Then
                sTiposDoc = sTiposDoc & IIf(Len(Trim(sTiposDoc)) > 0, SEPARADOR, "") & sTipoDoc
            End If
       
        Next
       
        If Len(Trim(sTiposDoc)) > 0 Then
       
            sTextoAux = sTiposDoc
            iPos = InStr(1, sTextoAux, SEPARADOR)
            Do While iPos <> 0
            
                Set objTitulos = New ClassTRPTitulos
                objTitulos.sTipoDoc = left(sTextoAux, iPos - 1)
                colTiposDoc.Add objTitulos
                
                sTextoAux = Mid(sTextoAux, iPos + 1)
                iPos = InStr(1, sTextoAux, SEPARADOR)
            
            Loop
            
            Set objTitulos = New ClassTRPTitulos
            objTitulos.sTipoDoc = sTextoAux
            colTiposDoc.Add objTitulos
       
        End If
        
        colCampos.Add "sTipoDoc"
        
        lErro = Ordena_Colecao(colTiposDoc, colSaida, colCampos)
        If lErro <> SUCESSO Then gError 200240
        
        sTipoDocAnt = ""
        sTipoDoc = ""
        For Each objTitulos In colSaida
            If sTipoDocAnt <> objTitulos.sTipoDoc Then
                sTipoDocAnt = objTitulos.sTipoDoc
                sTipoDoc = sTipoDoc & IIf(Len(Trim(sTipoDoc)) > 0, SEPARADOR, "") & objTitulos.sTipoDoc
            End If
        Next
        
        Set objTitulos = New ClassTRPTitulos
    
        objTitulos.iTipoDocOrigem = TRP_TIPO_DOC_DESTINO_TITPAG
        objTitulos.lNumIntDocOrigem = objTituloPagar.lNumIntDoc
        objTitulos.sTipoDoc = sTipoDoc
        
        lErro = CF("TRPTitulos_Insere", objTitulos, alComando(2))
        If lErro <> SUCESSO Then gError 200241
        
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    FaturaPagar_Grava_BD_Customizado = SUCESSO

    Exit Function

Erro_FaturaPagar_Grava_BD_Customizado:

    FaturaPagar_Grava_BD_Customizado = gErr

    Select Case gErr

        Case 200237
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200238, 200239
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRPTITULOS", gErr)
            
        Case 200240, 200241

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200242)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function
