VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTChequesPag3Mgz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get OptionFrente(ByVal objCT As Object) As Object
     Set OptionFrente = objCT.objUserControl.Controls("OptionFrente")
End Property

Public Property Get LabelNumCheque(ByVal objCT As Object) As Object
     Set LabelNumCheque = objCT.objUserControl.Controls("LabelNumCheque")
End Property

Public Property Get LabelValorCheque(ByVal objCT As Object) As Object
     Set LabelValorCheque = objCT.objUserControl.Controls("LabelValorCheque")
End Property

Public Property Get LabelBenefCheque(ByVal objCT As Object) As Object
     Set LabelBenefCheque = objCT.objUserControl.Controls("LabelBenefCheque")
End Property

Public Property Get TextoVerso(ByVal objCT As Object) As Object
     Set TextoVerso = objCT.objUserControl.Controls("TextoVerso")
End Property

Public Function CHQPAG3_ImprimirCheques(ByVal objCT As Object, lNumImpressao As Long, sLayoutCheques As String, dtDataEmissao As Date) As Long
'chama a impressao de cheques

Dim objRelatorio As New AdmRelatorio
Dim sNomeTsk As String
Dim lErro As Long, objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CHQPAG3_ImprimirCheques

    'a cidade deve vir do endereco da filial que está emitindo, se entrar como EMPRESA_TODA pegar da matriz
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then Error 19467
    
    lErro = objRelatorio.ExecutarDireto("Cheques", "", 0, sLayoutCheques, "NIMPRESSAO", CStr(lNumImpressao), "DEMISSAO", CStr(dtDataEmissao), "TCIDADE", objFilialEmpresa.objEndereco.sCidade, "TIGNORARMARGEM", "S", "TFRENTE", IIf(OptionFrente(objCT).Value, "S", "N"))
    If lErro <> SUCESSO Then Error 7431

    CHQPAG3_ImprimirCheques = SUCESSO

    Exit Function

Erro_CHQPAG3_ImprimirCheques:

    CHQPAG3_ImprimirCheques = Err

    Select Case Err

        Case 7431, 19467

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

Public Sub GridChequesPag3_RowColChange(ByVal objCT As Object)

Dim objInfoChequePag As ClassInfoChequePag

    Call Grid_RowColChange(objCT.objGridChequesPag3)
       
    If Not (objCT.gobjChequesPag Is Nothing) Then
    
        If objCT.GridChequesPag3.Row >= 1 And objCT.GridChequesPag3.Row <= objCT.objGridChequesPag3.iLinhasExistentes Then
            
            'Passa os dados da linha do Grid para o Obj
            Set objInfoChequePag = objCT.gobjChequesPag.ColInfoChequePag.Item(objCT.GridChequesPag3.Row)
            
            LabelNumCheque(objCT).Caption = CStr(objInfoChequePag.lNumRealCheque)
            LabelValorCheque(objCT).Caption = Format(objInfoChequePag.dValor, "Standard")
            LabelBenefCheque(objCT).Caption = objInfoChequePag.sFavorecido
            TextoVerso(objCT).Text = objInfoChequePag.sVerso
        
        End If
    
    End If

End Sub

Public Sub TextoVerso_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim objInfoChequePag As ClassInfoChequePag

    If Not (objCT.gobjChequesPag Is Nothing) Then
    
        If objCT.GridChequesPag3.Row >= 1 And objCT.GridChequesPag3.Row <= objCT.objGridChequesPag3.iLinhasExistentes Then
            
            'Passa os dados da linha do Grid para o Obj
            Set objInfoChequePag = objCT.gobjChequesPag.ColInfoChequePag.Item(objCT.GridChequesPag3.Row)
            
            objInfoChequePag.sVerso = TextoVerso(objCT).Text
        
        End If
    
    End If
    
End Sub
