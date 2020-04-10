VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoCOM"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Function Tipo_Requisitante(ByVal sRequisitante As String) As enumTipo

    If Len(Trim(sRequisitante)) = 0 Then
        Tipo_Requisitante = TIPO_VAZIO
    ElseIf Not IsNumeric(sRequisitante) Then
        Tipo_Requisitante = TIPO_STRING
    ElseIf Int(CDbl(sRequisitante)) <> CDbl(sRequisitante) Then
        Tipo_Requisitante = TIPO_DECIMAL
    ElseIf CDbl(sRequisitante) <= 0 Then
        Tipo_Requisitante = TIPO_NAO_POSITIVO
    ElseIf CDbl(sRequisitante) > NUM_MAX_VENDEDORES Then
        Tipo_Requisitante = TIPO_OVERFLOW
    Else
        Tipo_Requisitante = TIPO_CODIGO
    End If

End Function

Public Function TP_Requisitante_Le(objRequisitanteMaskEdBox As Object, objRequisitante As ClassRequisitante, Optional iCria As Integer = 1) As Long
'Lê o Requisitante com Código ou NomeRed em objRequisitanteMaskEdBox.Text
'Devolve em objRequisitante. Coloca NomeReduzido em .Text

Dim sRequisitante As String
Dim Requisitante As Object
Dim eTipoRequisitante As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Requisitante_Le

    Set Requisitante = objRequisitanteMaskEdBox
    sRequisitante = Trim(Requisitante.Text)
    eTipoRequisitante = Tipo_Requisitante(sRequisitante)

    Select Case eTipoRequisitante

        Case TIPO_STRING
    
            objRequisitante.sNomeReduzido = sRequisitante
            lErro = CF("Requisitante_Le_NomeReduzido", objRequisitante)
            If lErro <> SUCESSO And lErro <> 51152 Then Error 51142
            If lErro <> SUCESSO Then Error 51143
           
            Requisitante.Text = sRequisitante
    
        Case TIPO_CODIGO
    
            objRequisitante.lCodigo = CLng(sRequisitante)
            lErro = CF("Requisitante_Le", objRequisitante)
            If lErro <> SUCESSO And lErro <> 49084 Then Error 51144
            If lErro <> SUCESSO Then Error 51145
    
            Requisitante.Text = objRequisitante.sNomeReduzido
        
        Case TIPO_OVERFLOW
    
            Error 51146
    
        Case TIPO_DECIMAL
    
            Error 51147
    
        Case TIPO_NAO_POSITIVO
    
            Error 51148
    
    End Select

    TP_Requisitante_Le = SUCESSO

    Exit Function

TP_Requisitante_Le:

    TP_Requisitante_Le = Err

    Select Case Err
        
        Case 51142, 51144 'Tratados nas rotinas chamadas
        
            objRequisitanteMaskEdBox.SetFocus

        Case 51143  'Requisitante com NomeReduzido não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Requisitante não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_REQUISITANTE1", objRequisitante.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Requisitantes
                    lErro = Chama_Tela("Requisitante", objRequisitante)
                Else
                    'Segura o foco
                    objRequisitanteMaskEdBox.SetFocus
                End If
            Else
                'Segura o foco
                objRequisitanteMaskEdBox.SetFocus
            End If
        
        Case 51145  'Requisitante com Código não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Requisitante não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_REQUISITANTE", objRequisitante.lCodigo)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Requisitantes
                    lErro = Chama_Tela("Requisitante", objRequisitante)
                Else
                    'Segura o foco
                    objRequisitanteMaskEdBox.SetFocus
                End If
            Else
                'Segura o foco
                objRequisitanteMaskEdBox.SetFocus
            End If
        
        Case 51146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", Err, sRequisitante)
            objRequisitanteMaskEdBox.SetFocus

        Case 51147
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sRequisitante)
            objRequisitanteMaskEdBox.SetFocus

        Case 51148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sRequisitante)
            objRequisitanteMaskEdBox.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153609)

    End Select

End Function

'Function TP_Comprador_Le(objCompradorMaskEdBox As Object, objComprador As ClassComprador, Optional iCria As Integer = 1) As Long
''Lê o Comprador com Código ou NomeRed em objCompradorMaskEdBox.Text
''Devolve em objComprador. Coloca código-NomeReduzido no .Text
'
'Dim sComprador As String
'Dim iCodigo As Long
'Dim Comprador As object
'Dim lErro As Long
'Dim vbMsgRes As VbMsgBoxResult, X As New ClassChamaFuncoes
'Dim objUsuario As New ClassUsuario
'Dim objChamaFuncoes As New ClassChamaFuncoes
'
'On Error GoTo TP_Comprador_Le
'
'    Set Comprador = objCompradorMaskEdBox
'    sComprador = Trim(Comprador.Text)
'
'    'Tenta extrair código de sComprador
'    iCodigo = Codigo_Extrai(sComprador)
'
'    'Se é do tipo código
'    If iCodigo > 0 Then
'
'        objComprador.iCodigo = iCodigo
'        objComprador.iFilialEmpresa = giFilialEmpresa
'
'        'Lê o comprador
'        lErro = objChamaFuncoes.Comprador_Le(objComprador)
'        If lErro <> SUCESSO And lErro <> 50064 Then gError 68618
'        If lErro <> SUCESSO Then gError 68619
'
'        objUsuario.sCodUsuario = objComprador.sCodUsuario
'
'        'Lê o usuário
'        lErro = objChamaFuncoes.Usuario_Le(objUsuario)
'        If lErro <> SUCESSO And lErro <> 36347 Then gError 72550
'        If lErro <> SUCESSO Then gError 72551
'
'        Comprador.Text = CStr(objComprador.iCodigo) & SEPARADOR & objUsuario.sNomeReduzido
'
'    Else  'Se é do tipo String
'
'         objUsuario.sNomeReduzido = sComprador
'
'         lErro = objChamaFuncoes.Usuario_Le_NomeRed(objUsuario)
'         If lErro <> SUCESSO And lErro <> 57269 Then gError 68620
'         If lErro = 57269 Then gError 68621
'
'        objComprador.sCodUsuario = objUsuario.sCodUsuario
'
'        lErro = objChamaFuncoes.Comprador_Le_Usuario(objComprador)
'        If lErro <> SUCESSO And lErro <> 50059 Then gError 68622
'        If lErro <> SUCESSO Then gError 68623
'
'        Comprador.Text = CStr(objComprador.iCodigo) & SEPARADOR & sComprador
'
'
'    End If
'
'    TP_Comprador_Le = SUCESSO
'
'    Exit Function
'
'TP_Comprador_Le:
'
'    TP_Comprador_Le = gErr
'
'    Select Case gErr
'
'        Case 68618, 68620, 68622, 72550
'
'        Case 68619
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_NAO_CADASTRADO1", gErr, objComprador.sCodUsuario)
'
'        Case 68621
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO2", gErr, objUsuario.sNomeReduzido)
'
'        Case 68623
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR2", gErr, objUsuario.sCodUsuario)
'
'        Case 72551
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuario.sCodUsuario)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153610)
'
'    End Select
'
'End Function
'
