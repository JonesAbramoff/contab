VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoCRFAT"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Function Tipo_Cliente(ByVal sCliente As String) As enumTipo

If Len(Trim(sCliente)) = 0 Then
    Tipo_Cliente = TIPO_VAZIO
ElseIf Not IsNumeric(sCliente) Then
    Tipo_Cliente = TIPO_STRING
ElseIf Int(CDbl(sCliente)) <> CDbl(sCliente) Then
    Tipo_Cliente = TIPO_DECIMAL
ElseIf CDbl(sCliente) <= 0 Then
    Tipo_Cliente = TIPO_NAO_POSITIVO
ElseIf Len(Trim(sCliente)) > STRING_CGC Then
    Tipo_Cliente = TIPO_OVERFLOW
ElseIf Len(Trim(sCliente)) > STRING_CPF Then
    Tipo_Cliente = TIPO_CGC
ElseIf CDbl(sCliente) > NUM_MAX_CLIENTES Then
    Tipo_Cliente = TIPO_CPF
Else
    Tipo_Cliente = TIPO_CODIGO
End If

End Function

Private Function Tipo_Vendedor(ByVal sVendedor As String) As enumTipo

If Len(Trim(sVendedor)) = 0 Then
    Tipo_Vendedor = TIPO_VAZIO
ElseIf Not IsNumeric(sVendedor) Then
    Tipo_Vendedor = TIPO_STRING
ElseIf Int(CDbl(sVendedor)) <> CDbl(sVendedor) Then
    Tipo_Vendedor = TIPO_DECIMAL
ElseIf CDbl(sVendedor) <= 0 Then
    Tipo_Vendedor = TIPO_NAO_POSITIVO
ElseIf CDbl(sVendedor) > NUM_MAX_VENDEDORES Then
    Tipo_Vendedor = TIPO_OVERFLOW
Else
    Tipo_Vendedor = TIPO_CODIGO
End If

End Function

Public Function TP_Vendedor_Le(objVendedorMaskEdBox As Object, objVendedor As ClassVendedor, Optional iCria As Integer = 1) As Long
'Lê o Vendedor com Código ou NomeRed em objVendedorMaskEdBox.Text
'Devolve em objVendedor. Coloca NomeReduzido em .Text

Dim sVendedor As String
Dim Vendedor As Object
Dim eTipoVendedor As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Vendedor_Le

    Set Vendedor = objVendedorMaskEdBox
    sVendedor = Trim(Vendedor.Text)
    eTipoVendedor = Tipo_Vendedor(sVendedor)

    Select Case eTipoVendedor

    Case TIPO_STRING

        objVendedor.sNomeReduzido = sVendedor
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then Error 25010
        If lErro <> SUCESSO Then Error 25011
       
        Vendedor.Text = sVendedor

    Case TIPO_CODIGO

        objVendedor.iCodigo = CInt(sVendedor)
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then Error 25012
        If lErro <> SUCESSO Then Error 25013

        Vendedor.Text = objVendedor.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 25014

    Case TIPO_DECIMAL

        Error 25015

    Case TIPO_NAO_POSITIVO

        Error 25016

    End Select

    TP_Vendedor_Le = SUCESSO

    Exit Function

TP_Vendedor_Le:

    TP_Vendedor_Le = Err

    Select Case Err
        
        Case 25010, 25012 'Tratados nas rotinas chamadas
        
        Case 25011  'Vendedor com NomeReduzido não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Vendedor não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR2", objVendedor.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Vendedores
                    lErro = Chama_Tela("Vendedores", objVendedor)
                End If
            End If
        
        Case 25013  'Vendedor com Código não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Vendedor não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR1", objVendedor.iCodigo)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Vendedores
                    lErro = Chama_Tela("Vendedores", objVendedor)
                End If
            End If
        
        Case 25014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", Err, sVendedor)

        Case 25015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sVendedor)

        Case 25016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sVendedor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153616)

    End Select

End Function

Public Function TP_Vendedor_Le2(objVendedorMaskEdBox As Object, objVendedor As ClassVendedor, Optional iCria As Integer = 1) As Long
'Lê o Vendedor com Código ou NomeRed em objVendedorMaskEdBox.Text
'Devolve em objVendedor. Coloca código-NomeReduzido no .Text

Dim sVendedor As String
Dim iCodigo As Integer
Dim Vendedor As Object
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Vendedor_Le2

    Set Vendedor = objVendedorMaskEdBox
    sVendedor = Trim(Vendedor.Text)
    
    'Tenta extrair código de sVendedor
    iCodigo = Codigo_Extrai(sVendedor)
    
    'Se é do tipo código
    If iCodigo > 0 Then
    
        objVendedor.iCodigo = iCodigo
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then Error 25031
        If lErro <> SUCESSO Then Error 25032

        Vendedor.Text = CStr(objVendedor.iCodigo) & SEPARADOR & objVendedor.sNomeReduzido
        
    Else  'Se é do tipo String
            
         objVendedor.sNomeReduzido = sVendedor
         lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
         If lErro <> SUCESSO And lErro <> 25008 Then Error 25029
         If lErro <> SUCESSO Then Error 25030
        
         Vendedor.Text = CStr(objVendedor.iCodigo) & SEPARADOR & sVendedor
    
        
    End If

    TP_Vendedor_Le2 = SUCESSO

    Exit Function

TP_Vendedor_Le2:

    TP_Vendedor_Le2 = Err

    Select Case Err
        
        Case 25029, 25031 'Tratados nas rotinas chamadas

        Case 25030  'Vendedor com NomeReduzido não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Vendedor não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR2", objVendedor.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Vendedores
                    lErro = Chama_Tela("Vendedores", objVendedor)
                End If
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", Err, objVendedor.sNomeReduzido)
            End If
        
        Case 25032  'Vendedor com Código não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Vendedor não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR1", objVendedor.iCodigo)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Vendedores
                    lErro = Chama_Tela("Vendedores", objVendedor)
                End If
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", Err, objVendedor.iCodigo)
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153617)

    End Select

End Function

Public Function TP_Vendedor_Grid(objVendedorMaskEdBox As Object, objVendedor As ClassVendedor) As Long
'Lê o Vendedor com Código ou NomeRed em objVendedorMaskEdBox.Text
'Devolve em objVendedor. Para ser usada em Grids. Não trata foco nem chama tela.

Dim sVendedor As String
Dim Vendedor As Object
Dim eTipoVendedor As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Vendedor_Grid

    Set Vendedor = objVendedorMaskEdBox
    sVendedor = Trim(Vendedor.Text)
    eTipoVendedor = Tipo_Vendedor(sVendedor)

    Select Case eTipoVendedor

    Case TIPO_STRING

        objVendedor.sNomeReduzido = sVendedor
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then Error 25017
        If lErro <> SUCESSO Then Error 25018
       
        Vendedor.Text = sVendedor

    Case TIPO_CODIGO

        objVendedor.iCodigo = CInt(sVendedor)
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then Error 25019
        If lErro <> SUCESSO Then Error 25020

        Vendedor.Text = objVendedor.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 25021

    Case TIPO_DECIMAL

        Error 25022

    Case TIPO_NAO_POSITIVO

        Error 25023

    End Select

    TP_Vendedor_Grid = SUCESSO

    Exit Function

TP_Vendedor_Grid:

    TP_Vendedor_Grid = Err

    Select Case Err
        
        Case 25017, 25019 'Tratados nas rotinas chamadas

        Case 25018  'Vendedor com NomeReduzido não cadastrado
        
        Case 25020  'Vendedor com Código não cadastrado
        
        Case 25021
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", Err, sVendedor)

        Case 25022
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sVendedor)

        Case 25023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sVendedor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153618)

    End Select

End Function

Public Function TP_Cliente_Le(objClienteMaskEdBox As Object, objCliente As ClassCliente, iCodFilial As Integer, Optional iCria As Integer = 1, Optional iLoja As Integer = 0) As Long
'Lê o Cliente com Código ou NomeRed ou CGC/CPF em sCliente
'Devolve em objCliente e o código da filial em iCodFilial
'Será FILIAL_MATRIZ se sCliente não for CGC/CPF

Dim sCliente As String
Dim Cliente As Object
Dim eTipoCliente As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Cliente_Le

    Set Cliente = objClienteMaskEdBox
    sCliente = Trim(Cliente.Text)
    eTipoCliente = Tipo_Cliente(sCliente)

    Select Case eTipoCliente

    Case TIPO_STRING

        objCliente.sNomeReduzido = sCliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then Error 6666
        If lErro = 12348 Then Error 6668
       
        iCodFilial = FILIAL_MATRIZ
        Cliente.Text = sCliente

    Case TIPO_CODIGO

        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And iLoja = 1 Then
            objCliente.lCodigoLoja = StrParaLong(sCliente)
            objCliente.iFilialEmpresaLoja = giFilialEmpresa
        Else
            objCliente.lCodigo = StrParaLong(sCliente)
        End If
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then Error 6670
        If lErro = 12293 Then Error 6676

        iCodFilial = FILIAL_MATRIZ
        Cliente.Text = objCliente.sNomeReduzido

    Case TIPO_CPF

        'Critica o CPF
        lErro = Cpf_Critica(sCliente)
        If lErro <> SUCESSO Then Error 6677
        
        objCliente.sCgc = sCliente
        'Lê Cliente (com filial correspondente ao CGC) e código da Filial
        lErro = CF("Cliente_Le_Cgc", objCliente, iCodFilial)
        If lErro <> SUCESSO And lErro <> 6710 Then Error 6700
        If lErro = 6710 Then Error 6701

        Cliente.Text = objCliente.sNomeReduzido

    Case TIPO_CGC

        'Critica o CGC
        lErro = Cgc_Critica(sCliente)
        If lErro <> SUCESSO Then Error 6702
        
        objCliente.sCgc = sCliente
        'Lê Cliente (com filial correspondente ao CGC) e código da Filial
        lErro = CF("Cliente_Le_Cgc", objCliente, iCodFilial)
        If lErro <> SUCESSO And lErro <> 6710 Then Error 6703
        If lErro = 6710 Then Error 6704

        Cliente.Text = objCliente.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 6705

    Case TIPO_DECIMAL

        Error 6706

    Case TIPO_NAO_POSITIVO

        Error 6707

    End Select
    
    If objCliente.iAtivo <> CLIENTE_ATIVO Then
        Call Rotina_Aviso(vbOKOnly, "ERRO_CLIENTE_INATIVO", objCliente.lCodigo)
    End If

    TP_Cliente_Le = SUCESSO

    Exit Function

TP_Cliente_Le:

    TP_Cliente_Le = Err

    Select Case Err
                
        Case 6666, 6670, 6677, 6700, 6702, 6703

        Case 6668  'Cliente nao cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Cliente não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_1", objCliente.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Clientees
                    lErro = Chama_Tela("Clientes", objCliente)
                End If
                
            End If
        
        Case 6676  'Cliente nao cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Cliente não está cadastrado e pergunta se deseja criar
                If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_2", objCliente.lCodigoLoja)
                Else
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_2", objCliente.lCodigo)
                End If
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Clientees
                    lErro = Chama_Tela("Clientes", objCliente)
                End If
                
            End If
    
        Case 6701, 6704  'Cliente nao cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Cliente com este CGC/CPF não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_3", objCliente.sCgc)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Clientes
                    lErro = Chama_Tela("Clientes", objCliente)
                End If
            
            End If
        
        Case 6705
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_OVERFLOW", Err, sCliente)

        Case 6706
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sCliente)

        Case 6707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sCliente)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153619)

    End Select

    Exit Function

End Function

Function TP_Cliente_Le2(objClienteMaskEdBox As Object, objCliente As ClassCliente, Optional iCria As Integer = 1) As Long
'Lê o Cliente com Código ou NomeRed em objClienteMaskEdBox.Text
'Devolve em objCliente. Coloca código-NomeReduzido no .Text

Dim sCliente As String
Dim lCodigo As Long
Dim Cliente As Object
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Cliente_Le2

    Set Cliente = objClienteMaskEdBox
    sCliente = Trim(Cliente.Text)
    
    'Tenta extrair código de sCliente
    lCodigo = LCodigo_Extrai(sCliente)
    
    'Se é do tipo código
    If lCodigo > 0 Then
    
        objCliente.lCodigo = lCodigo
        
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then Error 47627
        If lErro <> SUCESSO Then Error 47628

        Cliente.Text = CStr(objCliente.lCodigo) & SEPARADOR & objCliente.sNomeReduzido
        
    Else  'Se é do tipo String
            
        objCliente.sNomeReduzido = sCliente
        
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then Error 47629
        If lErro <> SUCESSO Then Error 47630
        Cliente.Text = CStr(objCliente.lCodigo) & SEPARADOR & sCliente
    
    End If
    
    If objCliente.iAtivo <> CLIENTE_ATIVO Then
        Call Rotina_Aviso(vbOKOnly, "ERRO_CLIENTE_INATIVO", objCliente.lCodigo)
    End If
    
    TP_Cliente_Le2 = SUCESSO

    Exit Function

TP_Cliente_Le2:

    TP_Cliente_Le2 = Err

    Select Case Err
        
        Case 47627, 47629
            
        Case 47628  'Cliente com Código não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Cliente não está cadastrado e pergunta se deseja criar
                If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_2", objCliente.lCodigoLoja)
                Else
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_2", objCliente.lCodigo)
                End If
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Clientes
                    lErro = Chama_Tela("Clientes", objCliente)
                End If
            Else
                If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                    Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objCliente.lCodigoLoja)
                Else
                    Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objCliente.lCodigo)
                End If
            End If

        Case 47630  'Cliente com NomeReduzido não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Cliente não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_1", objCliente.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Clientes
                    lErro = Chama_Tela("Clientes", objCliente)
                End If
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, objCliente.sNomeReduzido)
            End If
        

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153620)

    End Select

End Function

Public Function TP_Cliente_Le3(objClienteMaskEdBox As Object, objCliente As ClassCliente, iCodFilial As Integer, Optional iCria As Integer = 1) As Long
'Lê o Cliente com Código ou NomeRed ou CGC/CPF em sCliente
'Devolve em objCliente e o código da filial em iCodFilial
'iCodFilial Será 0(zero) se sCliente não for CGC/CPF
'Quase igual a TP_Cliente_Le, só muda iCodFilial se sCliente não for CGC/CPF

Dim sCliente As String
Dim Cliente As Object
Dim eTipoCliente As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Cliente_Le3

    Set Cliente = objClienteMaskEdBox
    sCliente = Trim(Cliente.Text)
    eTipoCliente = Tipo_Cliente(sCliente)
    Cliente.ForeColor = vbWindowText

    Select Case eTipoCliente

    Case TIPO_STRING

        objCliente.sNomeReduzido = sCliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then Error 6666
        If lErro = 12348 Then Error 6668
       
        iCodFilial = 0
        Cliente.Text = sCliente

    Case TIPO_CODIGO

        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
            objCliente.lCodigoLoja = CLng(sCliente)
        Else
            objCliente.lCodigo = CLng(sCliente)
        End If
        
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then Error 6670
        If lErro = 12293 Then Error 6676

        iCodFilial = 0
        Cliente.Text = objCliente.sNomeReduzido

    Case TIPO_CPF

        'Critica o CPF
        lErro = Cpf_Critica(sCliente)
        If lErro <> SUCESSO Then Error 6677
        
        objCliente.sCgc = sCliente
        'Lê Cliente (com filial correspondente ao CGC) e código da Filial
        lErro = CF("Cliente_Le_Cgc", objCliente, iCodFilial)
        If lErro <> SUCESSO And lErro <> 6710 Then Error 6700
        If lErro = 6710 Then Error 6701

        Cliente.Text = objCliente.sNomeReduzido

    Case TIPO_CGC

        'Critica o CGC
        lErro = Cgc_Critica(sCliente)
        If lErro <> SUCESSO Then Error 6702
        
        objCliente.sCgc = sCliente
        'Lê Cliente (com filial correspondente ao CGC) e código da Filial
        lErro = CF("Cliente_Le_Cgc", objCliente, iCodFilial)
        If lErro <> SUCESSO And lErro <> 6710 Then Error 6703
        If lErro = 6710 Then Error 6704

        Cliente.Text = objCliente.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 6705

    Case TIPO_DECIMAL

        Error 6706

    Case TIPO_NAO_POSITIVO

        Error 6707

    End Select
    
    If objCliente.iAtivo <> CLIENTE_ATIVO Then
        Call Rotina_Aviso(vbOKOnly, "ERRO_CLIENTE_INATIVO", objCliente.lCodigo)
    End If
    
    TP_Cliente_Le3 = SUCESSO

    Exit Function

TP_Cliente_Le3:

    TP_Cliente_Le3 = Err

    Select Case Err
                
        Case 6666, 6670, 6677, 6700, 6702, 6703

        Case 6668  'Cliente nao cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Cliente não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_1", objCliente.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Clientees
                    lErro = Chama_Tela("Clientes", objCliente)
                End If
                
            End If
        
        Case 6676  'Cliente nao cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Cliente não está cadastrado e pergunta se deseja criar
                If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_2", objCliente.lCodigoLoja)
                Else
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_2", objCliente.lCodigo)
                End If
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Clientees
                    lErro = Chama_Tela("Clientes", objCliente)
                End If
                
            End If
    
        Case 6701, 6704  'Cliente nao cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Cliente com este CGC/CPF não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_3", objCliente.sCgc)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Clientes
                    lErro = Chama_Tela("Clientes", objCliente)
                End If
            
            End If
        
        Case 6705
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_OVERFLOW", Err, sCliente)

        Case 6706
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sCliente)

        Case 6707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sCliente)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153621)

    End Select

    Exit Function

End Function

Public Function TP_RegiaoVenda_Le(ByVal objMaskRegiaoVenda As Object, ByVal objRegiaoVenda As ClassRegiaoVenda, Optional ByVal iCria As Integer = 1) As Long
'recebe uma mask contendo, supostamente a descricao ou o codigo da regiao de venda...
'e faz a leitura da mesma.
'objMaskRegiaoVenda RECEBE(INPUT) o controle da tela que contém o conteúdo que servirá de base para a leitura
'objRegiaoVenda RETORNA(OUTPUT) os dados lidos do BD
'iCria RECEBE(INPUT) => a informação se é para tentar criar ou não a região, caso a mesma não seja encontrada

Dim lErro As Long
Dim eTipoPreenchimento As enumTipo

On Error GoTo Erro_TP_RegiaoVenda_Le
    
    'Obtem o tipo do preenchimento do campo RegiaoVenda.
    eTipoPreenchimento = Tipo_RegiaoVenda(objMaskRegiaoVenda.Text)
    
    'Seleciona pelo tipo obtido
    Select Case eTipoPreenchimento

        'Se for string, entao eh nomereduzido..
        Case TIPO_STRING
            
            'guarda regiaovenda.text no obj
            objRegiaoVenda.sDescricao = objMaskRegiaoVenda.Text
            
            'Le os dados da regiao
            lErro = CF("RegiaoVenda_Le_Descricao", objRegiaoVenda)
            If lErro <> SUCESSO And lErro <> 98828 Then gError 98827
            
            'se nao encontrou...
            If lErro <> SUCESSO Then gError 98829
                          
        'Se for codigo...
        Case TIPO_CODIGO
    
            'converte o conteudo da mascara para inteiro e armazena no obj
            objRegiaoVenda.iCodigo = StrParaInt(objMaskRegiaoVenda.Text)
            
            'Le os dados da regiao
            lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
            If lErro <> SUCESSO And lErro <> 16137 Then gError 98830
            
            ' se nao encontrou...
            If lErro <> SUCESSO Then gError 98831
            
        'Caso seja algo diferente dos anteriores,
        'erro...
        Case TIPO_OVERFLOW
    
            gError 98833
    
        Case TIPO_DECIMAL
    
            gError 98834
    
        Case TIPO_NAO_POSITIVO
    
            gError 98835

    End Select
   
    TP_RegiaoVenda_Le = SUCESSO
    
    Exit Function
    
Erro_TP_RegiaoVenda_Le:

    TP_RegiaoVenda_Le = gErr
    
    Select Case gErr
    
        Case 98827, 98830
    
        Case 98829
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA1", gErr, objRegiaoVenda.sDescricao)
    
        Case 98831
                
            If iCria = 1 Then
            
                'Envia aviso que Vendedor não está cadastrado e pergunta se deseja criar
                If Rotina_Aviso(vbYesNo, "AVISO_CRIAR_REGIAO") = vbYes Then
                    'Chama tela de Vendedores
                    lErro = Chama_Tela("RegiaoVenda", objRegiaoVenda)
                End If
            End If
    
        Case 98833
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", gErr, objMaskRegiaoVenda.Text)

        Case 98834
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", gErr, objMaskRegiaoVenda.Text)

        Case 98835
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, objMaskRegiaoVenda.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153622)
        
    End Select
    
    Exit Function

End Function

Public Function Tipo_RegiaoVenda(ByVal sRegiaoVenda As String) As enumTipo
'Funcao que verifica o tipo do parametro sRegiaoVenda

    'se regiao venda estiver vazio => tipo_vazio
    If Len(Trim(sRegiaoVenda)) = 0 Then
        Tipo_RegiaoVenda = TIPO_VAZIO
    
    'se regiao venda nao for numerico => tipo_string
    ElseIf Not IsNumeric(sRegiaoVenda) Then
        Tipo_RegiaoVenda = TIPO_STRING
    
    'se regiao venda convertido para dbl e depois para int for diferente de regiao venda convertido para dbl => tipo_decimal
    ElseIf Int(CDbl(sRegiaoVenda)) <> CDbl(sRegiaoVenda) Then
        Tipo_RegiaoVenda = TIPO_DECIMAL
    
    'se regiao venda, convertido para dbl for menor que 0 => tipo nao positivo
    ElseIf CDbl(sRegiaoVenda) <= 0 Then
        Tipo_RegiaoVenda = TIPO_NAO_POSITIVO
    
    'se regiao venda, convertido para dbl for maior que NUM_MAX_REGIAOVENDA => tipo overflow
    ElseIf CDbl(sRegiaoVenda) > NUM_MAX_REGIAOVENDA Then
        Tipo_RegiaoVenda = TIPO_OVERFLOW
    
    '... tipo codigo...
    Else
        Tipo_RegiaoVenda = TIPO_CODIGO
    
    End If

End Function

Public Function TP_Cliente_Le_Orcamento(objClienteMaskEdBox As Object, objCliente As ClassCliente, iCodFilial As Integer, Optional iCria As Integer = 1) As Long
'Lê o Cliente com Código ou NomeRed ou CGC/CPF em sCliente
'Devolve em objCliente e o código da filial em iCodFilial
'iCodFilial Será 0(zero) se sCliente não for CGC/CPF
'Quase igual a TP_Cliente_Le, só muda iCodFilial se sCliente não for CGC/CPF

Dim sCliente As String
Dim Cliente As Object
Dim eTipoCliente As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TP_Cliente_Le_Orcamento

    Set Cliente = objClienteMaskEdBox
    sCliente = Trim(Cliente.Text)
    eTipoCliente = Tipo_Cliente(sCliente)

    Select Case eTipoCliente

    Case TIPO_STRING

        objCliente.sNomeReduzido = sCliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 94435
        
        If lErro = 12348 Then 'gError 94436
       
            If iCria = 1 Then
            
                'Envia aviso que Cliente não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_TRIBUTACAO", objCliente.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    
                    'Chama tela de Clientees
                    lErro = Chama_Tela("Clientes", objCliente)
                    gError 94436
                
                Else
                    gError 94449
                    
                End If
                
            End If
       
        End If
       
        iCodFilial = 0
        Cliente.Text = sCliente

    Case TIPO_CODIGO

        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
            objCliente.lCodigoLoja = CLng(sCliente)
        Else
            objCliente.lCodigo = CLng(sCliente)
        End If
        
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 94437
        
        If lErro = 12293 Then 'gError 94438

            If iCria = 1 Then
            
                'Envia aviso que Cliente não está cadastrado e pergunta se deseja criar
                If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_TRIBUTACAO", objCliente.lCodigoLoja)
                Else
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_TRIBUTACAO", objCliente.lCodigo)
                End If
    
                If vbMsgRes = vbYes Then
                    
                    'Chama tela de Clientees
                    lErro = Chama_Tela("Clientes", objCliente)
                    gError 94438
                    
                Else
                    gError 94450
                    
                End If
                
            End If
        
        End If

        iCodFilial = 0
        Cliente.Text = objCliente.sNomeReduzido

    Case TIPO_CPF

        'Critica o CPF
        lErro = Cpf_Critica(sCliente)
        If lErro <> SUCESSO Then gError 94439
        
        objCliente.sCgc = sCliente
        'Lê Cliente (com filial correspondente ao CGC) e código da Filial
        lErro = CF("Cliente_Le_Cgc", objCliente, iCodFilial)
        If lErro <> SUCESSO And lErro <> 6710 Then gError 94440
        
        If lErro = 6710 Then 'gError 94441

            If iCria = 1 Then
            
                'Envia aviso que Cliente com este CGC/CPF não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_TRIBUTACAO", objCliente.sCgc)
    
                If vbMsgRes = vbYes Then
                    
                    'Chama tela de Clientes
                    lErro = Chama_Tela("Clientes", objCliente)
                    gError 94441
                    
                Else
                    gError 94451
                    
                End If
            
            End If
        
        End If
        
        Cliente.Text = objCliente.sNomeReduzido

    Case TIPO_CGC

        'Critica o CGC
        lErro = Cgc_Critica(sCliente)
        If lErro <> SUCESSO Then gError 94442
        
        objCliente.sCgc = sCliente
        'Lê Cliente (com filial correspondente ao CGC) e código da Filial
        lErro = CF("Cliente_Le_Cgc", objCliente, iCodFilial)
        If lErro <> SUCESSO And lErro <> 6710 Then gError 94443
        
        If lErro = 6710 Then 'gError 94444
        
            If iCria = 1 Then
            
                'Envia aviso que Cliente com este CGC/CPF não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_TRIBUTACAO", objCliente.sCgc)
    
                If vbMsgRes = vbYes Then
                    
                    'Chama tela de Clientes
                    lErro = Chama_Tela("Clientes", objCliente)
                    gError 94444
                    
                Else
                    gError 94452
                
                End If
            
            End If
        
        End If
        
        Cliente.Text = objCliente.sNomeReduzido
    
    Case TIPO_OVERFLOW

        gError 94445

    Case TIPO_DECIMAL

        gError 94446

    Case TIPO_NAO_POSITIVO

        gError 94447

    End Select

    TP_Cliente_Le_Orcamento = SUCESSO

    Exit Function

Erro_TP_Cliente_Le_Orcamento:

    TP_Cliente_Le_Orcamento = gErr

    Select Case gErr
                
        Case 94435, 94437, 94439, 94440, 94442, 94443

        Case 94436, 94438, 94441, 94444  'Cliente nao cadastrado e usuário resolveu cadastrar
                        
        Case 94449, 94450, 94451, 94452 'Cliente não cadastrado e o usuário não resolveu cadastrar
                        
        Case 94445
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_OVERFLOW", gErr, sCliente)

        Case 94446
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", gErr, sCliente)

        Case 94447
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sCliente)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153623)

    End Select

    Exit Function

End Function

'#############################################################################
'Inserido por Wagner 10/06/2006
Public Function TP_Contato_Le(objContatoMaskEdBox As Object, objContato As ClassContatos, iCodFilial As Integer, Optional iCria As Integer = 1, Optional iLoja As Integer = 0) As Long
'Lê o Contato com Código ou NomeRed ou CGC/CPF em sContato
'Devolve em objContato e o código da filial em iCodFilial
'Será FILIAL_MATRIZ se sContato não for CGC/CPF

Dim sContato As String
Dim Contato As Object
Dim eTipoContato As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Contato_Le

    Set Contato = objContatoMaskEdBox
    sContato = Trim(Contato.Text)
    eTipoContato = Tipo_Cliente(sContato)

    Select Case eTipoContato

    Case TIPO_STRING

        objContato.sNomeReduzido = sContato
        lErro = CF("Contato_Le_NomeReduzido", objContato)
        If lErro <> SUCESSO And lErro <> 180745 Then gError 181061
        If lErro = 180745 Then gError 181062
       
        iCodFilial = FILIAL_MATRIZ
        Contato.Text = sContato

    Case TIPO_CODIGO

        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And iLoja = 1 Then
            objContato.lCodigoLoja = StrParaLong(sContato)
            objContato.iFilialEmpresaLoja = giFilialEmpresa
        Else
            objContato.lCodigo = StrParaLong(sContato)
        End If
        lErro = CF("Contato_Le", objContato)
        If lErro <> SUCESSO And lErro <> 180548 Then gError 181063
        If lErro = 180548 Then gError 181064

        iCodFilial = FILIAL_MATRIZ
        Contato.Text = objContato.sNomeReduzido
    
    Case TIPO_OVERFLOW

        gError 181065

    Case TIPO_DECIMAL

        gError 181066

    Case TIPO_NAO_POSITIVO

        gError 181067

    End Select
    
    If objContato.iAtivo <> CLIENTE_ATIVO Then
        Call Rotina_Aviso(vbOKOnly, "ERRO_CLIENTE_INATIVO", objContato.lCodigo)
    End If

    TP_Contato_Le = SUCESSO

    Exit Function

TP_Contato_Le:

    TP_Contato_Le = gErr

    Select Case gErr
                
        Case 181061, 181063

        Case 181062  'Contato nao cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Contato não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONTATO", objContato.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Contatoes
                    Call Chama_Tela("Contatos", objContato)
                End If
                
            End If
        
        Case 181064  'Contato nao cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Contato não está cadastrado e pergunta se deseja criar
                If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONTATO", objContato.lCodigoLoja)
                Else
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONTATO", objContato.lCodigo)
                End If
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Contatoes
                    Call Chama_Tela("Contatos", objContato)
                End If
                
            End If
        
        Case 181065
            Call Rotina_Erro(vbOKOnly, "ERRO_CGC_OVERFLOW", gErr, sContato)

        Case 181066
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", gErr, sContato)

        Case 181067
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sContato)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181068)

    End Select

    Exit Function

End Function
'##################################################################################

Function TP_Contato_Le2(objContatoMaskEdBox As Object, objContato As ClassContatos, Optional iCria As Integer = 1) As Long
'Lê o Contato com Código ou NomeRed em objClienteMaskEdBox.Text
'Devolve em objCliente. Coloca código-NomeReduzido no .Text

Dim sContato As String
Dim lCodigo As Long
Dim Contato As Object
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Contato_Le2

    Set Contato = objContatoMaskEdBox
    sContato = Trim(Contato.Text)
    
    'Tenta extrair código de sContato
    lCodigo = LCodigo_Extrai(sContato)
    
    'Se é do tipo código
    If lCodigo > 0 Then
    
        objContato.lCodigo = lCodigo
        
        lErro = CF("Contato_Le", objContato)
        If lErro <> SUCESSO And lErro <> 180548 Then gError 193300
        If lErro <> SUCESSO Then gError 193301

        Contato.Text = CStr(objContato.lCodigo) & SEPARADOR & objContato.sNomeReduzido
        
    Else  'Se é do tipo String
            
        objContato.sNomeReduzido = sContato
        
        lErro = CF("Contato_Le_NomeReduzido", objContato)
        If lErro <> SUCESSO And lErro <> 180745 Then gError 193302
        
        If lErro <> SUCESSO Then gError 193303
        
        Contato.Text = CStr(objContato.lCodigo) & SEPARADOR & sContato
    
    End If
    
    If objContato.iAtivo <> CLIENTE_ATIVO Then
        Call Rotina_Aviso(vbOKOnly, "ERRO_CLIENTE_INATIVO", objContato.lCodigo)
    End If
    
    TP_Contato_Le2 = SUCESSO

    Exit Function

TP_Contato_Le2:

    TP_Contato_Le2 = gErr

    Select Case gErr
        
        Case 193300, 193302
        
        Case 193301

            If iCria = 1 Then
            
                'Envia aviso que Contato não está cadastrado e pergunta se deseja criar
                If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONTATO", objContato.lCodigoLoja)
                Else
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONTATO", objContato.lCodigo)
                End If
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Contatoes
                    Call Chama_Tela("Contatos", objContato)
                End If
                
            Else
                If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
                    Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_NAO_CADASTRADO", gErr, objContato.lCodigoLoja)
                Else
                    Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_NAO_CADASTRADO", gErr, objContato.lCodigo)
                End If
                
            End If



        Case 193303  'Contato com NomeReduzido não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Contato não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONTATO", objContato.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Contatoes
                    Call Chama_Tela("Contatos", objContato)
                End If
                
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_NAO_CADASTRADO1", gErr, objContato.sNomeReduzido)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193304)

    End Select

End Function

