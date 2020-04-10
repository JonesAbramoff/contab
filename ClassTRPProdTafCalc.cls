VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPProdTafCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolTarifas As New Collection
Private mvarsProduto As String
Private mvarcolTarifaAcordo As New Collection

Public Property Set colTarifaAcordo(ByVal vData As Collection)
    Set mvarcolTarifaAcordo = vData
End Property

Public Property Get colTarifaAcordo() As Collection
    Set colTarifaAcordo = mvarcolTarifaAcordo
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Set colTarifas(ByVal vData As Collection)
    Set mvarcolTarifas = vData
End Property

Public Property Get colTarifas() As Collection
    Set colTarifas = mvarcolTarifas
End Property

Public Function Calcula_Tarifa(ByVal dtData As Date, ByVal iVigencia As Integer, dValor As Double) As Long
'Calcula a tarifa unitária de folheto

Dim objTaf As ClassTRPProdTarifa
Dim objTafAcor As ClassTRPAcordoTarifaDif
Dim dTarifa As Double
Dim dtDataVigencia As Date
Dim iIndice As Integer

On Error GoTo Erro_Calcula_Tarifa

    dtDataVigencia = DATA_NULA
    For Each objTaf In mvarcolTarifas
        If objTaf.dtVigencia > dtDataVigencia And objTaf.dtVigencia <= dtData Then
            dtDataVigencia = objTaf.dtVigencia
        End If
    Next
    
    If colTarifaAcordo.Count = 0 Then
        For iIndice = mvarcolTarifas.Count To 1 Step -1
            
            Set objTaf = mvarcolTarifas.Item(iIndice)
            
            If objTaf.iDiasDe <= iVigencia And (objTaf.iDiasAte >= iVigencia Or objTaf.iDiario = MARCADO) And objTaf.dtVigencia = dtDataVigencia Then
            
                If objTaf.iDiasAte < iVigencia Then
                    dTarifa = objTaf.dValor + (objTaf.dValorAdicional * (iVigencia - objTaf.iDiasAte))
                Else
                    dTarifa = objTaf.dValor
                End If
            
                Exit For
            End If
        
        Next
    Else
        
        For iIndice = mvarcolTarifaAcordo.Count To 1 Step -1
            
            Set objTafAcor = mvarcolTarifaAcordo.Item(iIndice)
            
            If objTafAcor.iDiasDe <= iVigencia And (objTafAcor.iDiasAte >= iVigencia Or objTafAcor.iDiario = MARCADO) Then
            
                If objTafAcor.iDiasAte < iVigencia Then
                    dTarifa = objTafAcor.dValor + (objTafAcor.dValorAdicional * (iVigencia - objTafAcor.iDiasAte))
                Else
                    dTarifa = objTafAcor.dValor
                End If
            
                Exit For
            End If
        
        Next
    End If
    
    dValor = dTarifa

    Calcula_Tarifa = SUCESSO

    Exit Function

Erro_Calcula_Tarifa:

    Calcula_Tarifa = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192205)

    End Select

    Exit Function
    
End Function

