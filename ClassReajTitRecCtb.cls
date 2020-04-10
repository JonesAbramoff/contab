VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassReajTitRecCtb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CCL1 As String = "Ccl"

Private mvarobjFormCTBAux As Object

Private gobjContabAutomatica As ClassContabAutomatica
Private mvarProcReajusteTitRec As ClassProcReajTitRec

Public Property Set objProcReajusteTitRec(ByVal vData As ClassProcReajTitRec)
    Set mvarProcReajusteTitRec = vData
End Property

Public Property Get objProcReajusteTitRec() As ClassProcReajTitRec
    Set objProcReajusteTitRec = mvarProcReajusteTitRec
End Property

Public Property Get Controls() As Object
    
    If mvarobjFormCTBAux Is Nothing Then
    
        Set mvarobjFormCTBAux = New FormAuxCTB
        
    End If
    
    Set Controls = mvarobjFormCTBAux.Controls
    
End Property

Private Sub Class_Terminate()
    Set mvarobjFormCTBAux = Nothing
End Sub

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é responsavel por gerar a contabilizacao

Dim lErro As Long, lDoc As Long, iConta As Integer, dValorLivroAux As Double

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    
    'obtem numero de doc para a filial
    lErro = objContabAutomatica.Obter_Doc(lDoc, mvarProcReajusteTitRec.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 130199
    
    Controls("GridCcl").Tag = mvarProcReajusteTitRec.colCcl.Count
        
    'grava a contabilizacao
    lErro = objContabAutomatica.Gravar_Registro(Me, "ReajusteTitRec", objProcReajusteTitRec.lNumIntDoc, 0, 0, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, mvarProcReajusteTitRec.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 130200
            
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = gErr
     
    Select Case gErr
          
        Case 130199, 130200
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152782)
     
    End Select
     
    Exit Function

End Function

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, objMnemonico As New ClassMnemonicoCTBValor
Dim sCclMascarado As String, dValorReajusteTotal As Double
Dim objProcReajTitRecCcl As ClassProcReajTitRecCcl

On Error GoTo Erro_Calcula_Mnemonico

        Select Case objMnemonicoValor.sMnemonico
    
            Case "ValorReajuste_Ccl"
                For Each objProcReajTitRecCcl In mvarProcReajusteTitRec.colCcl
                    
                    objMnemonicoValor.colValor.Add objProcReajTitRecCcl.dValorReajuste
                    
                Next
            
            Case "ValorReajuste_Total"
                For Each objProcReajTitRecCcl In mvarProcReajusteTitRec.colCcl
                    
                    dValorReajusteTotal = Arredonda_Moeda(dValorReajusteTotal + objProcReajTitRecCcl.dValorReajuste)
                    
                Next
                objMnemonicoValor.colValor.Add dValorReajusteTotal
            
            Case CCL1
            
                For Each objProcReajTitRecCcl In mvarProcReajusteTitRec.colCcl
                    
                    lErro = Mascara_MascararCcl(objProcReajTitRecCcl.sCcl, sCclMascarado)
                    If lErro <> SUCESSO Then gError 30241
                    
                    objMnemonicoValor.colValor.Add sCclMascarado
                    
                Next
            
            Case Else
                Error 39695
    
        End Select
    
    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err
        
        Case 39695
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152783)

    End Select

    Exit Function

End Function

