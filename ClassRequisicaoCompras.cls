VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRequisicaoCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvardtData As Date
Private mvardtDataEnvio As Date
Private mvardtDataLimite As Date
Private mvarlUrgente As Long
Private mvarlRequisitante As Long
Private mvarsCCL As String
Private mvarsOPCodigo As String
Private mvarlPVCodigo As Long
Private mvariFilialCompra As Integer
Private mvariTipoDestino As Integer
Private mvarlFornCliDestino As Long
Private mvariFilialDestino As Integer
Private mvarlObservacao As Long
Private mvarcolItens As New Collection
Private mvarsObservacao As String
Private mvardtDataBaixa As Date
Private mvariTipoTributacao As Integer
Private mvariSelecionado As Integer
Private mvarsDigitador As String
Private mvarsNomeFilEmp As String
Private mvarsNomeRequisitante As String
Private mvarsNomeFilialCompra As String
Private mvarlNumIntDocItemOP As Long
Private mvarobjTela As Object
Private mvarbInclusao As Boolean
Private mvarbEnvio As Boolean
Private mvarbAprovacao As Boolean
Private mvarbBaixa As Boolean
Private mvarbRecebimento As Boolean
Private mvarsUsuReg As String
Private mvarsUsuRegEnvio As String
Private mvarsUsuRegAprov As String
Private mvarsUsuRegBaixa As String
Private mvarName As String

Public Property Let Name(ByVal vData As String)
    mvarName = vData
End Property

Public Property Get Name() As String
    Name = mvarName
End Property

Public Property Let sUsuRegBaixa(ByVal vData As String)
    mvarsUsuRegBaixa = vData
End Property

Public Property Get sUsuRegBaixa() As String
    sUsuRegBaixa = mvarsUsuRegBaixa
End Property

Public Property Let sUsuRegAprov(ByVal vData As String)
    mvarsUsuRegAprov = vData
End Property

Public Property Get sUsuRegAprov() As String
    sUsuRegAprov = mvarsUsuRegAprov
End Property

Public Property Let sUsuRegEnvio(ByVal vData As String)
    mvarsUsuRegEnvio = vData
End Property

Public Property Get sUsuRegEnvio() As String
    sUsuRegEnvio = mvarsUsuRegEnvio
End Property

Public Property Let sUsuReg(ByVal vData As String)
    mvarsUsuReg = vData
End Property

Public Property Get sUsuReg() As String
    sUsuReg = mvarsUsuReg
End Property

Public Property Let bRecebimento(ByVal vData As Boolean)
    mvarbRecebimento = vData
End Property

Public Property Get bRecebimento() As Boolean
    bRecebimento = mvarbRecebimento
End Property

Public Property Let bBaixa(ByVal vData As Boolean)
    mvarbBaixa = vData
End Property

Public Property Get bBaixa() As Boolean
    bBaixa = mvarbBaixa
End Property

Public Property Let bAprovacao(ByVal vData As Boolean)
    mvarbAprovacao = vData
End Property

Public Property Get bAprovacao() As Boolean
    bAprovacao = mvarbAprovacao
End Property

Public Property Let bEnvio(ByVal vData As Boolean)
    mvarbEnvio = vData
End Property

Public Property Get bEnvio() As Boolean
    bEnvio = mvarbEnvio
End Property

Public Property Let bInclusao(ByVal vData As Boolean)
    mvarbInclusao = vData
End Property

Public Property Get bInclusao() As Boolean
    bInclusao = mvarbInclusao
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Let lNumIntDocItemOP(ByVal vData As Long)
    mvarlNumIntDocItemOP = vData
End Property

Public Property Get lNumIntDocItemOP() As Long
    lNumIntDocItemOP = mvarlNumIntDocItemOP
End Property

Public Property Let sNomeFilialCompra(ByVal vData As String)
    mvarsNomeFilialCompra = vData
End Property

Public Property Get sNomeFilialCompra() As String
    sNomeFilialCompra = mvarsNomeFilialCompra
End Property

Public Property Let sNomeRequisitante(ByVal vData As String)
    mvarsNomeRequisitante = vData
End Property

Public Property Get sNomeRequisitante() As String
    sNomeRequisitante = mvarsNomeRequisitante
End Property

Public Property Let sNomeFilEmp(ByVal vData As String)
    mvarsNomeFilEmp = vData
End Property

Public Property Get sNomeFilEmp() As String
    sNomeFilEmp = mvarsNomeFilEmp
End Property

Public Property Let sDigitador(ByVal vData As String)
Attribute sDigitador.VB_Description = "Código do usuário digitador."
    mvarsDigitador = vData
End Property


Public Property Get sDigitador() As String
    sDigitador = mvarsDigitador
End Property



Public Property Let iSelecionado(ByVal vData As Integer)
    mvariSelecionado = vData
End Property


Public Property Get iSelecionado() As Integer
    iSelecionado = mvariSelecionado
End Property



Public Property Let iTipoTributacao(ByVal vData As Integer)
Attribute iTipoTributacao.VB_Description = "Tipo de Tributação"
    mvariTipoTributacao = vData
End Property


Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property



Public Property Let dtDataBaixa(ByVal vData As Date)
    mvardtDataBaixa = vData
End Property


Public Property Get dtDataBaixa() As Date
    dtDataBaixa = mvardtDataBaixa
End Property


Public Property Let sObservacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sObservacao = 5
    mvarsObservacao = vData
End Property


Public Property Get sObservacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sObservacao
    sObservacao = mvarsObservacao
End Property




Public Property Set colItens(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colItens = Form1
    Set mvarcolItens = vData
End Property


Public Property Get colItens() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colItens
    Set colItens = mvarcolItens
End Property



Public Property Let lObservacao(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lObservacao = 5
    mvarlObservacao = vData
End Property


Public Property Get lObservacao() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lObservacao
    lObservacao = mvarlObservacao
End Property



Public Property Let iFilialDestino(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialDestino = 5
    mvariFilialDestino = vData
End Property


Public Property Get iFilialDestino() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialDestino
    iFilialDestino = mvariFilialDestino
End Property



Public Property Let lFornCliDestino(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lFornCliDestino = 5
    mvarlFornCliDestino = vData
End Property


Public Property Get lFornCliDestino() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lFornCliDestino
    lFornCliDestino = mvarlFornCliDestino
End Property



Public Property Let iTipoDestino(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoDestino = 5
    mvariTipoDestino = vData
End Property


Public Property Get iTipoDestino() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoDestino
    iTipoDestino = mvariTipoDestino
End Property



Public Property Let iFilialCompra(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialCompra = 5
    mvariFilialCompra = vData
End Property


Public Property Get iFilialCompra() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialCompra
    iFilialCompra = mvariFilialCompra
End Property



Public Property Let lPVCodigo(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lPVCodigo = 5
    mvarlPVCodigo = vData
End Property


Public Property Get lPVCodigo() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lPVCodigo
    lPVCodigo = mvarlPVCodigo
End Property



Public Property Let sOPCodigo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sOPCodigo = 5
    mvarsOPCodigo = vData
End Property


Public Property Get sOPCodigo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sOPCodigo
    sOPCodigo = mvarsOPCodigo
End Property



Public Property Let sCCL(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCcl = 5
    mvarsCCL = vData
End Property


Public Property Get sCCL() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCcl
    sCCL = mvarsCCL
End Property



Public Property Let lRequisitante(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lRequisitante = 5
    mvarlRequisitante = vData
End Property


Public Property Get lRequisitante() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lRequisitante
    lRequisitante = mvarlRequisitante
End Property



Public Property Let lUrgente(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lUrgente = 5
    mvarlUrgente = vData
End Property


Public Property Get lUrgente() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lUrgente
    lUrgente = mvarlUrgente
End Property



Public Property Let dtDataLimite(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataLimite = 5
    mvardtDataLimite = vData
End Property


Public Property Get dtDataLimite() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataLimite
    dtDataLimite = mvardtDataLimite
End Property



Public Property Let dtDataEnvio(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataEnvio = 5
    mvardtDataEnvio = vData
End Property


Public Property Get dtDataEnvio() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataEnvio
    dtDataEnvio = mvardtDataEnvio
End Property



Public Property Let dtData(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtData = 5
    mvardtData = vData
End Property


Public Property Get dtData() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtData
    dtData = mvardtData
End Property



Public Property Let lCodigo(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCodigo = 5
    mvarlCodigo = vData
End Property


Public Property Get lCodigo() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCodigo
    lCodigo = mvarlCodigo
End Property



Public Property Let iFilialEmpresa(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialEmpresa = 5
    mvariFilialEmpresa = vData
End Property


Public Property Get iFilialEmpresa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialEmpresa
    iFilialEmpresa = mvariFilialEmpresa
End Property



Public Property Let lNumIntDoc(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDoc = 5
    mvarlNumIntDoc = vData
End Property


Public Property Get lNumIntDoc() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDoc
    lNumIntDoc = mvarlNumIntDoc
End Property

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor, Optional objContexto As Object) As Long

Dim lErro As Long
Dim sRequisitante As String
Dim objRequisitante As New ClassRequisitante
Dim objUsu As ClassUsuarios, sNome As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
                        
        Case "CodigoRequisicao"
            objMnemonicoValor.colValor.Add lCodigo
            
        Case "FilialEmpresa"
            objMnemonicoValor.colValor.Add iFilialEmpresa
            
        Case "Inclusao_ReqCompra"
            objMnemonicoValor.colValor.Add bInclusao
            
        Case "Envio_ReqCompra"
            objMnemonicoValor.colValor.Add bEnvio
            
        Case "Aprovacao_ReqCompra"
            objMnemonicoValor.colValor.Add bAprovacao
        
        Case "Baixa_ReqCompra"
            objMnemonicoValor.colValor.Add bBaixa
                       
        Case "Requisitante"
        
            sRequisitante = ""
            If lRequisitante <> 0 Then
                
                objRequisitante.lCodigo = lRequisitante
        
                'Le o requisitante para colocar o NomeReduzido na tela
                lErro = CF("Requisitante_Le", objRequisitante)
                If lErro <> SUCESSO And lErro <> 49084 Then gError ERRO_SEM_MENSAGEM
            
                sRequisitante = objRequisitante.sNomeReduzido
            
            End If

            objMnemonicoValor.colValor.Add sRequisitante
            
        Case "UsuCadastro"
                
            objMnemonicoValor.colValor.Add sUsuReg
            
        Case "Aprovador"
            objMnemonicoValor.colValor.Add sUsuRegAprov
            
        Case "UsuEnvio"
            objMnemonicoValor.colValor.Add sUsuRegEnvio
             
        Case "UsuBaixa"
            objMnemonicoValor.colValor.Add sUsuRegBaixa
            
        Case "EmailRequisitante"
        
            sEmail = ""
            If lRequisitante <> 0 Then
                
                objRequisitante.lCodigo = lRequisitante
        
                'Le o requisitante para colocar o NomeReduzido na tela
                lErro = CF("Requisitante_Le", objRequisitante)
                If lErro <> SUCESSO And lErro <> 49084 Then gError ERRO_SEM_MENSAGEM
            
                sEmail = objRequisitante.sEmail
                
                If sEmail = "" Then
                    Set objUsu = New ClassUsuarios
                    objUsu.sCodUsuario = objRequisitante.sCodUsuario
                    If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                        lErro = CF("Usuarios_Le", objUsu)
                        If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                        sEmail = objUsu.sEmail
                    End If
                End If
            End If

            objMnemonicoValor.colValor.Add sEmail
            
        Case "EmailUsuCadastro"
        
            sEmail = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuReg
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sEmail = objUsu.sEmail
            End If

            objMnemonicoValor.colValor.Add sEmail
            
        Case "EmailAprovador"

            sEmail = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuRegAprov
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sEmail = objUsu.sEmail
            End If
            
            objMnemonicoValor.colValor.Add sEmail
            
        Case "EmailUsuEnvio"
        
            sEmail = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuEnvio
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sEmail = objUsu.sEmail
            End If
            
            objMnemonicoValor.colValor.Add sEmail
             
        Case "EmailUsuBaixa"
        
            sEmail = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuRegBaixa
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sEmail = objUsu.sEmail
            End If
            
            objMnemonicoValor.colValor.Add sEmail
            
        Case "NomeUsuCadastro"
        
            sNome = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuReg
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sNome = objUsu.sNome
            End If

            objMnemonicoValor.colValor.Add sNome
            
        Case "NomeAprovador"

            sNome = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuRegAprov
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sNome = objUsu.sNome
            End If
            
            objMnemonicoValor.colValor.Add sNome
            
        Case "NomeUsuEnvio"
        
            sNome = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuEnvio
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sNome = objUsu.sNome
            End If
            
            objMnemonicoValor.colValor.Add sNome
             
        Case "NomeUsuBaixa"
        
            sNome = ""
            Set objUsu = New ClassUsuarios
            objUsu.sCodUsuario = sUsuRegBaixa
            If Len(Trim(objUsu.sCodUsuario)) > 0 Then
                lErro = CF("Usuarios_Le", objUsu)
                If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
                sNome = objUsu.sNome
            End If
            
            objMnemonicoValor.colValor.Add sNome
              
        Case Else
            gError 178683

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 178683
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178684)

    End Select

    Exit Function

End Function

Private Sub Class_Initialize()
    mvarbAprovacao = False
    mvarbBaixa = False
    mvarbEnvio = False
    mvarbInclusao = False
End Sub
