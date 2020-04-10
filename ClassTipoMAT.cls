VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Function Tipo_Campo(ByVal sCampo As String) As enumTipo

If Len(Trim(sCampo)) = 0 Then
    Tipo_Campo = TIPO_VAZIO
ElseIf Not IsNumeric(sCampo) Then
    Tipo_Campo = TIPO_STRING
ElseIf Int(CDbl(sCampo)) <> CDbl(sCampo) Then
    Tipo_Campo = TIPO_DECIMAL
ElseIf CDbl(sCampo) <= 0 Then
    Tipo_Campo = TIPO_NAO_POSITIVO
ElseIf CDbl(sCampo) > NUM_MAX_ALMOXARIFADOS Then
    Tipo_Campo = TIPO_OVERFLOW
Else
    Tipo_Campo = TIPO_CODIGO
End If

End Function

'Lê Almoxarifado com Código ou NomeRed em objAlmoxarifadoMaskEdBox.Text
'Devolve em objAlmoxarifado. Coloca NomeReduzido em .Text
Public Function TP_Almoxarifado_Le(objAlmoxarifadoMaskEdBox As Object, objAlmoxarifado As ClassAlmoxarifado, Optional iCria As Integer = 1) As Long

Dim sAlmoxarifado As String
Dim Almoxarifado As Object
Dim eTipoAlmoxarifado As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult, iCodigo As Integer

On Error GoTo TP_Almoxarifado_Le

    Set Almoxarifado = objAlmoxarifadoMaskEdBox
    sAlmoxarifado = Trim(Almoxarifado.Text)
    
    'incluido p/tratar codigo-nomereduzido
    iCodigo = Codigo_Extrai(sAlmoxarifado)
    If iCodigo <> 0 Then
        sAlmoxarifado = CStr(iCodigo)
    End If
        
    eTipoAlmoxarifado = Tipo_Campo(sAlmoxarifado)

    Select Case eTipoAlmoxarifado

    Case TIPO_STRING

        objAlmoxarifado.sNomeReduzido = sAlmoxarifado
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then Error 25010
        If lErro <> SUCESSO Then Error 25011
        
        Almoxarifado.Text = objAlmoxarifado.sNomeReduzido

    Case TIPO_CODIGO

        objAlmoxarifado.iCodigo = CInt(sAlmoxarifado)
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then Error 25012
        If lErro <> SUCESSO Then Error 25013

        Almoxarifado.Text = objAlmoxarifado.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 25014

    Case TIPO_DECIMAL

        Error 25015

    Case TIPO_NAO_POSITIVO

        Error 25016

    End Select

    TP_Almoxarifado_Le = SUCESSO

    Exit Function

TP_Almoxarifado_Le:

    TP_Almoxarifado_Le = Err

    Select Case Err
        
        Case 25010, 25012 'Tratados nas rotinas chamadas
        
            Almoxarifado.SetFocus

        Case 25011  'Almoxarifado com NomeReduzido não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Almoxarifado não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", objAlmoxarifado.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Almoxarifadoes
                    lErro = Chama_Tela("Almoxarifado", objAlmoxarifado)
                Else
                    'Segura o foco
                    Almoxarifado.SetFocus
                End If
            Else
                'Segura o foco
                Almoxarifado.SetFocus
            End If
        
        Case 25013  'Almoxarifado com Código não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Almoxarifado não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO1", objAlmoxarifado.iCodigo)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Almoxarifados
                    lErro = Chama_Tela("Almoxarifado", objAlmoxarifado)
                Else
                    'Segura o foco
                    Almoxarifado.SetFocus
                End If
            Else
                'Segura o foco
                Almoxarifado.SetFocus
            End If
        
        Case 25014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", Err, sAlmoxarifado)
            Almoxarifado.SetFocus

        Case 25015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sAlmoxarifado)
            Almoxarifado.SetFocus

        Case 25016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sAlmoxarifado)
            Almoxarifado.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153629)

    End Select

End Function

Public Function TP_Almoxarifado_Filial_Le(objAlmoxarifadoMaskEdBox As Object, objAlmoxarifado As ClassAlmoxarifado, Optional iCria As Integer = 1) As Long
'Lê Almoxarifado de giFilialEmpresa com Código ou NomeRed em objAlmoxarifadoMaskEdBox.Text
'Devolve em objAlmoxarifado. Coloca NomeReduzido em .Text

    TP_Almoxarifado_Filial_Le = CF("TP_Almoxarifado_Filial_Le_Aux", objAlmoxarifadoMaskEdBox, objAlmoxarifado, iCria)

End Function
'Lê Almoxarifado com Código ou NomeRed em objAlmoxarifadoMaskEdBox.Text
'Devolve em objAlmoxarifado. Coloca NomeReduzido em .Text
'Para ser usada em Grids. Não segura foco nem chama tela.
Public Function TP_Almoxarifado_Grid(objAlmoxarifadoMaskEdBox As Object, objAlmoxarifado As ClassAlmoxarifado) As Long

Dim sAlmoxarifado As String
Dim Almoxarifado As Object
Dim eTipoAlmoxarifado As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Almoxarifado_Grid

    Set Almoxarifado = objAlmoxarifadoMaskEdBox
    sAlmoxarifado = Trim(Almoxarifado.Text)
    eTipoAlmoxarifado = Tipo_Campo(sAlmoxarifado)

    Select Case eTipoAlmoxarifado

    Case TIPO_STRING

        objAlmoxarifado.sNomeReduzido = sAlmoxarifado
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then Error 25113
        If lErro <> SUCESSO Then Error 25114
        
        Almoxarifado.Text = sAlmoxarifado

    Case TIPO_CODIGO

        objAlmoxarifado.iCodigo = CInt(sAlmoxarifado)
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then Error 25116
        If lErro <> SUCESSO Then Error 25117

        Almoxarifado.Text = objAlmoxarifado.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 25119

    Case TIPO_DECIMAL

        Error 25120

    Case TIPO_NAO_POSITIVO

        Error 25121

    End Select

    TP_Almoxarifado_Grid = SUCESSO

    Exit Function

TP_Almoxarifado_Grid:

    TP_Almoxarifado_Grid = Err

    Select Case Err
        
        Case 25113, 25116 'Tratados nas rotinas chamadas
        
        Case 25114  'Almoxarifado com NomeReduzido não cadastrado
        
        Case 25117  'Almoxarifado com Código não cadastrado
         
        Case 25119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", Err, sAlmoxarifado)

        Case 25120
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sAlmoxarifado)

        Case 25121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sAlmoxarifado)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153630)

    End Select

End Function

Public Function TP_Almoxarifado_Filial_Grid(objAlmoxarifadoMaskEdBox As Object, objAlmoxarifado As ClassAlmoxarifado) As Long
'Lê Almoxarifado de giFilialEmpresa com Código ou NomeRed em objAlmoxarifadoMaskEdBox.Text
'Devolve em objAlmoxarifado. Coloca NomeReduzido em .Text
'Para ser usada em Grids. Não segura foco nem chama tela.

    TP_Almoxarifado_Filial_Grid = CF("TP_Almoxarifado_Filial_Grid_Aux", objAlmoxarifadoMaskEdBox, objAlmoxarifado)

End Function

'Lê Almoxarifado com Código ou NomeRed em objAlmoxarifadoMaskEdBox.Text
'Critica se Almoxarifado tem o produto com código sProduto cadastrado.
'Devolve em objAlmoxarifado. Coloca NomeReduzido em .Text
'Para ser usada em Grids. Não segura foco nem chama tela.
Public Function TP_Almoxarifado_Produto_Grid(sProduto As String, objAlmoxarifadoMaskEdBox As Object, objAlmoxarifado As ClassAlmoxarifado) As Long

Dim sAlmoxarifado As String
Dim Almoxarifado As Object
Dim eTipoAlmoxarifado As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim colAlmoxarifado As New Collection
Dim vAlmoxarifado As Variant
Dim iTemProduto As Integer

On Error GoTo TP_Almoxarifado_Produto_Grid

    Set Almoxarifado = objAlmoxarifadoMaskEdBox
    sAlmoxarifado = Trim(Almoxarifado.Text)
    eTipoAlmoxarifado = Tipo_Campo(sAlmoxarifado)

    Select Case eTipoAlmoxarifado

    Case TIPO_STRING

        objAlmoxarifado.sNomeReduzido = sAlmoxarifado
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then Error 25129
        If lErro <> SUCESSO Then Error 25130
        
        'Lê Almoxarifados que tem sProduto cadastrado
        lErro = CF("AlmoxarifadosProduto_Le", sProduto, colAlmoxarifado)
        If lErro <> SUCESSO Then Error 25132
       
        'Verifica se objAlmoxarifado tem o Produto
        iTemProduto = 0
        For Each vAlmoxarifado In colAlmoxarifado
        
            If vAlmoxarifado = objAlmoxarifado.iCodigo Then
                iTemProduto = 1
                Exit For
            End If
        
        Next
        
        'Se não tem, erro
        If iTemProduto = 0 Then Error 25133
        
        Almoxarifado.Text = sAlmoxarifado

    Case TIPO_CODIGO

        objAlmoxarifado.iCodigo = CInt(sAlmoxarifado)
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then Error 25134
        If lErro <> SUCESSO Then Error 25135

        'Lê Almoxarifados que tem sProduto cadastrado
        lErro = CF("AlmoxarifadosProduto_Le", sProduto, colAlmoxarifado)
        If lErro <> SUCESSO Then Error 25137
       
        'Verifica se objAlmoxarifado tem o Produto
        iTemProduto = 0
        For Each vAlmoxarifado In colAlmoxarifado
        
            If vAlmoxarifado = objAlmoxarifado.iCodigo Then
                iTemProduto = 1
                Exit For
            End If
        
        Next
        
        'Se não tem, erro
        If iTemProduto = 0 Then Error 25138
        
        Almoxarifado.Text = objAlmoxarifado.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 25139

    Case TIPO_DECIMAL

        Error 25140

    Case TIPO_NAO_POSITIVO

        Error 25141

    End Select

    TP_Almoxarifado_Produto_Grid = SUCESSO

    Exit Function

TP_Almoxarifado_Produto_Grid:

    TP_Almoxarifado_Produto_Grid = Err

    Select Case Err
        
        Case 25129, 25134 'Tratados nas rotinas chamadas
        
        Case 25130  'Almoxarifado com NomeReduzido não cadastrado
       
        Case 25132, 25137 'Tratados na rotina chamada
        
        Case 25133, 25138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_TEM_PRODUTO", Err, objAlmoxarifado.iCodigo, sProduto)

        Case 25135  'Almoxarifado com Código não cadastrado
         
        Case 25139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", Err, sAlmoxarifado)

        Case 25140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sAlmoxarifado)

        Case 25141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sAlmoxarifado)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153631)

    End Select

End Function

Public Function TP_Almoxarifado_Filial_Produto_Grid(sProduto As String, objAlmoxarifadoMaskEdBox As Object, objAlmoxarifado As ClassAlmoxarifado) As Long
'Lê Almoxarifado de giFilialEmpresa com Código ou NomeRed em objAlmoxarifadoMaskEdBox.Text
'Critica se Almoxarifado tem o produto com código sProduto cadastrado.
'Devolve em objAlmoxarifado. Coloca NomeReduzido em .Text
'Para ser usada em Grids. Não segura foco nem chama tela.

    TP_Almoxarifado_Filial_Produto_Grid = CF("TP_Almoxarifado_Filial_Produto_Grid_Aux", sProduto, objAlmoxarifadoMaskEdBox, objAlmoxarifado)

End Function

'Lê Almoxarifado com Código, NomeRed ou Codigo-NomeRed em objAlmoxarifadoMaskEdBox.Text
'Devolve em objAlmoxarifado. Coloca Codigo-NomeReduzido em .Text
Public Function TP_Almoxarifado_Le_ComCodigo(objAlmoxarifadoMaskEdBox As Object, objAlmoxarifado As ClassAlmoxarifado, Optional iCria As Integer = 0) As Long

Dim sAlmoxarifado As String, iCodigo As Integer
Dim Almoxarifado As Object
Dim eTipoAlmoxarifado As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Almoxarifado_Le_ComCodigo

    Set Almoxarifado = objAlmoxarifadoMaskEdBox
    sAlmoxarifado = Trim(Almoxarifado.Text)
    
    'incluido p/tratar codigo-nomereduzido
    iCodigo = Codigo_Extrai(sAlmoxarifado)
    If iCodigo <> 0 Then
        sAlmoxarifado = CStr(iCodigo)
    End If
    
    eTipoAlmoxarifado = Tipo_Campo(sAlmoxarifado)

    Select Case eTipoAlmoxarifado

    Case TIPO_STRING

        objAlmoxarifado.sNomeReduzido = sAlmoxarifado
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then Error 27779
        If lErro <> SUCESSO Then Error 27780
        
        Almoxarifado.Text = CStr(objAlmoxarifado.iCodigo) & SEPARADOR & sAlmoxarifado

    Case TIPO_CODIGO

        objAlmoxarifado.iCodigo = iCodigo
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then Error 27781
        If lErro <> SUCESSO Then Error 27782

        Almoxarifado.Text = sAlmoxarifado & SEPARADOR & objAlmoxarifado.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 27783

    Case TIPO_DECIMAL

        Error 27784

    Case TIPO_NAO_POSITIVO

        Error 27785

    End Select

    TP_Almoxarifado_Le_ComCodigo = SUCESSO

    Exit Function

TP_Almoxarifado_Le_ComCodigo:

    TP_Almoxarifado_Le_ComCodigo = Err

    Select Case Err
        
        Case 27779, 27781 'Tratados nas rotinas chamadas
        
            Almoxarifado.SetFocus

        Case 27780  'Almoxarifado com NomeReduzido não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Almoxarifado não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", objAlmoxarifado.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Almoxarifadoes
                    lErro = Chama_Tela("Almoxarifado", objAlmoxarifado)
                Else
                    'Segura o foco
                    Almoxarifado.SetFocus
                End If
            Else
                
                lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO1", Err, objAlmoxarifado.sNomeReduzido)
                
                'Segura o foco
                Almoxarifado.SetFocus
            End If
        
        Case 27782  'Almoxarifado com Código não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Almoxarifado não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO1", objAlmoxarifado.iCodigo)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Almoxarifados
                    lErro = Chama_Tela("Almoxarifado", objAlmoxarifado)
                Else
                    'Segura o foco
                    Almoxarifado.SetFocus
                End If
            Else
                
                lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", Err, objAlmoxarifado.iCodigo)
                
                'Segura o foco
                Almoxarifado.SetFocus
            End If
        
        Case 27783
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", Err, sAlmoxarifado)
            Almoxarifado.SetFocus

        Case 27784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sAlmoxarifado)
            Almoxarifado.SetFocus

        Case 27785
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sAlmoxarifado)
            Almoxarifado.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153632)

    End Select

End Function

Public Function TP_Almoxarifado_Filial_Le_ComCodigo(objAlmoxarifadoMaskEdBox As Object, objAlmoxarifado As ClassAlmoxarifado, Optional iCria As Integer = 0) As Long
'Lê Almoxarifado com Código, NomeRed ou Codigo-NomeRed em objAlmoxarifadoMaskEdBox.Text
'Devolve em objAlmoxarifado. Coloca Codigo-NomeReduzido em .Text

    TP_Almoxarifado_Filial_Le_ComCodigo = CF("TP_Almoxarifado_Filial_Le_ComCodigo_Aux", objAlmoxarifadoMaskEdBox, objAlmoxarifado, iCria)

End Function

Public Function TP_FilialEmpresa_Le(sFilialEmpresa As String, objFilialEmpresa As AdmFiliais) As Long
'Lê FilialEmpresa com Código ou NomeRed e coloca seu resultado em objFilialEmpresa

Dim eTipo As enumTipo
Dim lErro As Long
Dim iAchou As Integer
Dim iFilialEmpresa As Integer

On Error GoTo TP_FilialEmpresa_Le

    iAchou = 0

    'incluido p/tratar codigo-nomereduzido
    iFilialEmpresa = Codigo_Extrai(sFilialEmpresa)
    If iFilialEmpresa <> 0 Then
        sFilialEmpresa = CStr(iFilialEmpresa)
    End If

    eTipo = Tipo_Campo(sFilialEmpresa)

    Select Case eTipo

        Case TIPO_STRING
    
            For Each objFilialEmpresa In gcolFiliais
                If objFilialEmpresa.sNome = sFilialEmpresa Then
                    iAchou = 1
                    Exit For
                End If
            Next
    
            If iAchou = 0 Then gError 71971
    
        Case TIPO_CODIGO
    
            For Each objFilialEmpresa In gcolFiliais
                If objFilialEmpresa.iCodFilial = iFilialEmpresa Then
                    iAchou = 1
                    Exit For
                End If
            Next
    
            If iAchou = 0 Then gError 71972
    
        Case TIPO_OVERFLOW
    
            gError 71973
    
        Case TIPO_DECIMAL
    
            gError 71974
    
        Case TIPO_NAO_POSITIVO
    
            gError 71975

    End Select

    TP_FilialEmpresa_Le = SUCESSO

    Exit Function

TP_FilialEmpresa_Le:

    TP_FilialEmpresa_Le = gErr

    Select Case gErr
        
        Case 71971, 71972 'Filial não existente
        
        Case 71973
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", gErr, sFilialEmpresa)

        Case 71974
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", gErr, sFilialEmpresa)

        Case 71975
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sFilialEmpresa)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153633)

    End Select

End Function

