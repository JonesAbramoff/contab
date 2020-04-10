VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoCPR"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Function Tipo_PrestServ(ByVal sPrestServ As String) As enumTipo

If Len(Trim(sPrestServ)) = 0 Then
    Tipo_PrestServ = TIPO_VAZIO
ElseIf Not IsNumeric(sPrestServ) Then
    Tipo_PrestServ = TIPO_STRING
ElseIf Int(CDbl(sPrestServ)) <> CDbl(sPrestServ) Then
    Tipo_PrestServ = TIPO_DECIMAL
ElseIf CDbl(sPrestServ) <= 0 Then
    Tipo_PrestServ = TIPO_NAO_POSITIVO
ElseIf CDbl(sPrestServ) > NUM_MAX_FORNECEDORES Then
    Tipo_PrestServ = TIPO_OVERFLOW
Else
    Tipo_PrestServ = TIPO_CODIGO
End If

End Function

Public Function TP_PrestServ_Le(objPrestServMaskEdBox As Object, objPrestServ As ClassPrestServ, Optional iCria As Integer = 1) As Long
'Lê o PrestServ com Código ou NomeRed em objPrestServMaskEdBox.Text
'Devolve em objPrestServ. Coloca NomeReduzido em .Text

Dim sPrestServ As String
Dim PrestServ As Object
Dim eTipoPrestServ As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_PrestServ_Le

    Set PrestServ = objPrestServMaskEdBox
    sPrestServ = Trim(PrestServ.Text)
    eTipoPrestServ = Tipo_PrestServ(sPrestServ)

    Select Case eTipoPrestServ

    Case TIPO_STRING

        objPrestServ.sNomeReduzido = sPrestServ
        lErro = CF("PrestServ_Le_NomeReduzido", objPrestServ)
        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then Error 25010
        If lErro <> SUCESSO Then Error 25011
       
        PrestServ.Text = sPrestServ

    Case TIPO_CODIGO

        objPrestServ.lCodigo = CLng(sPrestServ)
        lErro = CF("PrestServ_Le", objPrestServ)
        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then Error 25012
        If lErro <> SUCESSO Then Error 25013

        PrestServ.Text = objPrestServ.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 25014

    Case TIPO_DECIMAL

        Error 25015

    Case TIPO_NAO_POSITIVO

        Error 25016

    End Select

    TP_PrestServ_Le = SUCESSO

    Exit Function

TP_PrestServ_Le:

    TP_PrestServ_Le = Err

    Select Case Err
        
        Case 25010, 25012 'Tratados nas rotinas chamadas
        
        Case 25011  'PrestServ com NomeReduzido não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que PrestServ não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRESTSERV2", objPrestServ.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de PrestServ
                    lErro = Chama_Tela("PrestServ", objPrestServ)
                End If
            End If
        
        Case 25013  'PrestServ com Código não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que PrestServ não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRESTSERV1", objPrestServ.lCodigo)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de PrestServ
                    lErro = Chama_Tela("PrestServ", objPrestServ)
                End If
                
            End If
        
        Case 25014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", Err, sPrestServ)

        Case 25015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sPrestServ)

        Case 25016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sPrestServ)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153611)

    End Select

End Function

Private Function Tipo_Fornecedor(ByVal sFornecedor As String) As enumTipo

If Len(Trim(sFornecedor)) = 0 Then
    Tipo_Fornecedor = TIPO_VAZIO
ElseIf Not IsNumeric(sFornecedor) Then
    Tipo_Fornecedor = TIPO_STRING
ElseIf Int(CDbl(sFornecedor)) <> CDbl(sFornecedor) Then
    Tipo_Fornecedor = TIPO_DECIMAL
ElseIf CDbl(sFornecedor) <= 0 Then
    Tipo_Fornecedor = TIPO_NAO_POSITIVO
ElseIf Len(Trim(sFornecedor)) > STRING_CGC Then
    Tipo_Fornecedor = TIPO_OVERFLOW
ElseIf Len(Trim(sFornecedor)) > STRING_CPF Then
    Tipo_Fornecedor = TIPO_CGC
ElseIf CDbl(sFornecedor) > NUM_MAX_FORNECEDORES Then
    Tipo_Fornecedor = TIPO_CPF
Else
    Tipo_Fornecedor = TIPO_CODIGO
End If

End Function

Public Function TP_Fornecedor_Le(objFornecedorMaskEdBox As Object, objFornecedor As ClassFornecedor, iCodFilial As Integer, Optional iCria As Integer = 1) As Long
'Lê o Fornecedor com Código ou NomeRed ou CGC/CPF em sFornecedor
'Devolve em objFornecedor e o código da filial em iCodFilial
'Será FILIAL_MATRIZ se sFornecedor não for CGC/CPF

Dim sFornecedor As String
Dim Fornecedor As Object
Dim eTipoFornecedor As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Fornecedor_Le

    Set Fornecedor = objFornecedorMaskEdBox
    sFornecedor = Trim(Fornecedor.Text)
    eTipoFornecedor = Tipo_Fornecedor(sFornecedor)

    Select Case eTipoFornecedor

    Case TIPO_STRING

        objFornecedor.sNomeReduzido = sFornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 6655
        If lErro = 6681 Then Error 6663
       
        iCodFilial = FILIAL_MATRIZ
        Fornecedor.Text = sFornecedor

    Case TIPO_CODIGO

        objFornecedor.lCodigo = CLng(sFornecedor)
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then Error 6659
        If lErro = 12729 Then Error 6664

        iCodFilial = FILIAL_MATRIZ
        Fornecedor.Text = objFornecedor.sNomeReduzido

    Case TIPO_CPF

        'Critica o CPF
        lErro = Cpf_Critica(sFornecedor)
        If lErro <> SUCESSO Then Error 6667
        
        objFornecedor.sCgc = sFornecedor
        'Lê Fornecedor (com filial correspondente ao CGC) e código da Filial
        lErro = CF("Fornecedor_Le_Cgc", objFornecedor, iCodFilial)
        If lErro <> SUCESSO And lErro <> 6694 Then Error 6674
        If lErro = 6694 Then Error 6675

        Fornecedor.Text = objFornecedor.sNomeReduzido

    Case TIPO_CGC

        'Critica o CGC
        lErro = Cgc_Critica(sFornecedor)
        If lErro <> SUCESSO Then Error 6669
        
        objFornecedor.sCgc = sFornecedor
        'Lê Fornecedor (com filial correspondente ao CGC) e código da Filial
        lErro = CF("Fornecedor_Le_Cgc", objFornecedor, iCodFilial)
        If lErro <> SUCESSO And lErro <> 6694 Then Error 6656
        If lErro = 6694 Then Error 6660

        Fornecedor.Text = objFornecedor.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 6671

    Case TIPO_DECIMAL

        Error 6672

    Case TIPO_NAO_POSITIVO

        Error 6673

    End Select

    If objFornecedor.iAtivo = 0 Then
        Call Rotina_Aviso(vbOKOnly, "ERRO_FORNECEDOR_INATIVO", objFornecedor.lCodigo)
    End If

    TP_Fornecedor_Le = SUCESSO

    Exit Function

TP_Fornecedor_Le:

    TP_Fornecedor_Le = Err

    Select Case Err
        
        Case 6655, 6656, 6659, 6667, 6669, 6674
        
            Fornecedor.SetFocus

        Case 6663  'Fornecedor não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Fornecedor não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_1", objFornecedor.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Fornecedores
                    lErro = Chama_Tela("Fornecedores", objFornecedor)
                Else
                    'Segura o foco
                    Fornecedor.SetFocus
                End If
            Else
                'Segura o foco
                Fornecedor.SetFocus
            End If
        
        Case 6664  'Fornecedor nao cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Fornecedor não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_2", objFornecedor.lCodigo)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Fornecedores
                    lErro = Chama_Tela("Fornecedores", objFornecedor)
                Else
                    'Segura o foco
                    Fornecedor.SetFocus
                End If
            Else
                'Segura o foco
                Fornecedor.SetFocus
            End If
        
        Case 6660, 6675  'Fornecedor nao cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Fornecedor com este CGC/CPF não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_3", objFornecedor.sCgc)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Fornecedores
                    lErro = Chama_Tela("Fornecedores", objFornecedor)
                Else
                    'Segura o foco
                    Fornecedor.SetFocus
                End If
            Else
                'Segura o foco
                Fornecedor.SetFocus
            End If
        
        Case 6671
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_OVERFLOW", Err, sFornecedor)
            Fornecedor.SetFocus

        Case 6672
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sFornecedor)
            Fornecedor.SetFocus

        Case 6673
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sFornecedor)
            Fornecedor.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153612)

    End Select

    Exit Function

End Function

Public Function TP_Fornecedor_Grid(objFornecedorMaskEdBox As Object, objFornecedor As ClassFornecedor, iCodFilial As Integer) As Long
'Lê o Fornecedor com Código ou NomeRed ou CGC/CPF em sFornecedor
'Devolve em objFornecedor e o código da filial em iCodFilial
'Será FILIAL_MATRIZ se sFornecedor não for CGC/CPF.
'Para ser usada em GRID. Não chama tela. Não seta foco.

Dim sFornecedor As String
Dim Fornecedor As Object
Dim eTipoFornecedor As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Fornecedor_Grid

    Set Fornecedor = objFornecedorMaskEdBox
    sFornecedor = Trim(Fornecedor.Text)
    eTipoFornecedor = Tipo_Fornecedor(sFornecedor)

    Select Case eTipoFornecedor

    Case TIPO_STRING

        objFornecedor.sNomeReduzido = sFornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 25610
        If lErro = 6681 Then Error 25611
       
        'Devolve iCodFilial não preenchido
        iCodFilial = 0
        Fornecedor.Text = sFornecedor

    Case TIPO_CODIGO

        objFornecedor.lCodigo = CLng(sFornecedor)
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then Error 25612
        If lErro = 12729 Then Error 25613

        'Devolve iCodFilial não preenchido
        iCodFilial = 0
        Fornecedor.Text = objFornecedor.sNomeReduzido

    Case TIPO_CPF

        'Critica o CPF
        lErro = Cpf_Critica(sFornecedor)
        If lErro <> SUCESSO Then Error 25614
        
        objFornecedor.sCgc = sFornecedor
        'Lê Fornecedor (com filial correspondente ao CGC) e código da Filial
        lErro = CF("Fornecedor_Le_Cgc", objFornecedor, iCodFilial)
        If lErro <> SUCESSO And lErro <> 6694 Then Error 25615
        If lErro = 6694 Then Error 25616

        Fornecedor.Text = objFornecedor.sNomeReduzido

    Case TIPO_CGC

        'Critica o CGC
        lErro = Cgc_Critica(sFornecedor)
        If lErro <> SUCESSO Then Error 25617
        
        objFornecedor.sCgc = sFornecedor
        'Lê Fornecedor (com filial correspondente ao CGC) e código da Filial
        lErro = CF("Fornecedor_Le_Cgc", objFornecedor, iCodFilial)
        If lErro <> SUCESSO And lErro <> 6694 Then Error 25618
        If lErro = 6694 Then Error 25619

        Fornecedor.Text = objFornecedor.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 25620

    Case TIPO_DECIMAL

        Error 25621

    Case TIPO_NAO_POSITIVO

        Error 25622

    End Select

    If objFornecedor.iAtivo = 0 Then
        Call Rotina_Aviso(vbOKOnly, "ERRO_FORNECEDOR_INATIVO", objFornecedor.lCodigo)
    End If

    TP_Fornecedor_Grid = SUCESSO

    Exit Function

TP_Fornecedor_Grid:

    TP_Fornecedor_Grid = Err

    Select Case Err
        
        Case 25610, 25612, 25614, 25615, 25617, 25618
            'Tratado na rotina chamada
        
        Case 25611  'Fornecedor com NomeReduzido não cadastrado
        
        Case 25613  'Fornecedor com código não cadastrado.
        
        Case 25616, 25619 'Fornecedor com CGC/CPF nao cadastrado.
        
        Case 25620
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_OVERFLOW", Err, sFornecedor)

        Case 25621
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sFornecedor)

        Case 25622
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sFornecedor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153613)

    End Select

    Exit Function

End Function

Function TP_Fornecedor_Le2(objFornecedorMaskEdBox As Object, objFornecedor As ClassFornecedor, Optional iCria As Integer = 1) As Long
'Lê o Fornecedor com Código ou NomeRed em objFornecedorMaskEdBox.Text
'Devolve em objFornecedor. Coloca código-NomeReduzido no .Text

Dim sFornecedor As String
Dim lCodigo As Long
Dim Fornecedor As Object
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Fornecedor_Le2

    Set Fornecedor = objFornecedorMaskEdBox
    sFornecedor = Trim(Fornecedor.Text)
    
    'Tenta extrair código de sFornecedor
    lCodigo = LCodigo_Extrai(sFornecedor)
    
    'Se é do tipo código
    If lCodigo > 0 Then
    
        objFornecedor.lCodigo = lCodigo
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then Error 47665
        If lErro <> SUCESSO Then Error 47666

        Fornecedor.Text = CStr(objFornecedor.lCodigo) & SEPARADOR & objFornecedor.sNomeReduzido
        
    Else  'Se é do tipo String
            
         objFornecedor.sNomeReduzido = sFornecedor
         
         lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
         If lErro <> SUCESSO And lErro <> 6681 Then Error 47667
         If lErro <> SUCESSO Then Error 47668
        
         Fornecedor.Text = CStr(objFornecedor.lCodigo) & SEPARADOR & sFornecedor
        
    End If

    If objFornecedor.iAtivo = 0 Then
        Call Rotina_Aviso(vbOKOnly, "ERRO_FORNECEDOR_INATIVO", objFornecedor.lCodigo)
    End If

    TP_Fornecedor_Le2 = SUCESSO

    Exit Function

TP_Fornecedor_Le2:

    TP_Fornecedor_Le2 = Err

    Select Case Err
        
        Case 47665, 47667
            
        Case 47666  'Fornecedor com Código não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Fornecedor não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_2", objFornecedor.lCodigo)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Fornecedors
                    lErro = Chama_Tela("Fornecedores", objFornecedor)
                Else
                    Fornecedor.SetFocus
                End If
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", Err, lCodigo)
                Fornecedor.SetFocus
            End If

        Case 47668  'Fornecedor com NomeReduzido não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Fornecedor não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_1", objFornecedor.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Fornecedors
                    lErro = Chama_Tela("Fornecedores", objFornecedor)
                Else
                    Fornecedor.SetFocus
                End If
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, sFornecedor)
                Fornecedor.SetFocus
            End If
        

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153614)

    End Select

End Function

Public Function TP_Fornecedor_Le3(objFornecedorMaskEdBox As Object, objFornecedor As ClassFornecedor, iCodFilial As Integer, Optional iCria As Integer = 1) As Long
'Lê o Fornecedor com Código ou NomeRed ou CGC/CPF em sFornecedor
'Devolve em objFornecedor e o código da filial em iCodFilial
'iCodFilial Será 0(zero) se Fornecedor não for CGC/CPF
'Quase igual a TP_Fornecedor_Le, só muda iCodFilial se Fornecedor não for CGC/CPF

Dim sFornecedor As String
Dim Fornecedor As Object
Dim eTipoFornecedor As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Fornecedor_Le3

    Set Fornecedor = objFornecedorMaskEdBox
    sFornecedor = Trim(Fornecedor.Text)
    eTipoFornecedor = Tipo_Fornecedor(sFornecedor)

    Select Case eTipoFornecedor

    Case TIPO_STRING

        objFornecedor.sNomeReduzido = sFornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 6655
        If lErro = 6681 Then Error 6663
       
        iCodFilial = 0
        Fornecedor.Text = sFornecedor

    Case TIPO_CODIGO

        objFornecedor.lCodigo = CLng(sFornecedor)
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then Error 6659
        If lErro = 12729 Then Error 6664

        iCodFilial = 0
        Fornecedor.Text = objFornecedor.sNomeReduzido

    Case TIPO_CPF

        'Critica o CPF
        lErro = Cpf_Critica(sFornecedor)
        If lErro <> SUCESSO Then Error 6667
        
        objFornecedor.sCgc = sFornecedor
        'Lê Fornecedor (com filial correspondente ao CGC) e código da Filial
        lErro = CF("Fornecedor_Le_Cgc", objFornecedor, iCodFilial)
        If lErro <> SUCESSO And lErro <> 6694 Then Error 6674
        If lErro = 6694 Then Error 6675

        Fornecedor.Text = objFornecedor.sNomeReduzido

    Case TIPO_CGC

        'Critica o CGC
        lErro = Cgc_Critica(sFornecedor)
        If lErro <> SUCESSO Then Error 6669
        
        objFornecedor.sCgc = sFornecedor
        'Lê Fornecedor (com filial correspondente ao CGC) e código da Filial
        lErro = CF("Fornecedor_Le_Cgc", objFornecedor, iCodFilial)
        If lErro <> SUCESSO And lErro <> 6694 Then Error 6656
        If lErro = 6694 Then Error 6660

        Fornecedor.Text = objFornecedor.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 6671

    Case TIPO_DECIMAL

        Error 6672

    Case TIPO_NAO_POSITIVO

        Error 6673

    End Select

    If objFornecedor.iAtivo = 0 Then
        Call Rotina_Aviso(vbOKOnly, "ERRO_FORNECEDOR_INATIVO", objFornecedor.lCodigo)
    End If

    TP_Fornecedor_Le3 = SUCESSO

    Exit Function

TP_Fornecedor_Le3:

    TP_Fornecedor_Le3 = Err

    Select Case Err
        
        Case 6655, 6656, 6659, 6667, 6669, 6674
        
            Fornecedor.SetFocus

        Case 6663  'Fornecedor não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Fornecedor não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_1", objFornecedor.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Fornecedores
                    lErro = Chama_Tela("Fornecedores", objFornecedor)
                Else
                    'Segura o foco
                    Fornecedor.SetFocus
                End If
            Else
                'Segura o foco
                Fornecedor.SetFocus
            End If
        
        Case 6664  'Fornecedor nao cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Fornecedor não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_2", objFornecedor.lCodigo)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Fornecedores
                    lErro = Chama_Tela("Fornecedores", objFornecedor)
                Else
                    'Segura o foco
                    Fornecedor.SetFocus
                End If
            Else
                'Segura o foco
                Fornecedor.SetFocus
            End If
        
        Case 6660, 6675  'Fornecedor nao cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Fornecedor com este CGC/CPF não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FORNECEDOR_3", objFornecedor.sCgc)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Fornecedores
                    lErro = Chama_Tela("Fornecedores", objFornecedor)
                Else
                    'Segura o foco
                    Fornecedor.SetFocus
                End If
            Else
                'Segura o foco
                Fornecedor.SetFocus
            End If
        
        Case 6671
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_OVERFLOW", Err, sFornecedor)
            Fornecedor.SetFocus

        Case 6672
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sFornecedor)
            Fornecedor.SetFocus

        Case 6673
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sFornecedor)
            Fornecedor.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153615)

    End Select

    Exit Function

End Function


