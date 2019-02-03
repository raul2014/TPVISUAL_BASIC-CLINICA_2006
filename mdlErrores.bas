Attribute VB_Name = "mdlErrores"
Option Explicit

Public Enum AccionError
    Ignorar = 1
    Reintentar = 2
    Finalizar = 3
    Cancelar = 4
End Enum

Public Function evaluarError(ByVal Err As ErrObject) As AccionError
    Dim res As Variant
    Select Case Err.Number
        Case 1
        Case 2
        Case 3
       
       'Case n
                    
       'division por 0
       
            
        'no coinciden los tipos o subindice fuera de intervalo
        Case 13, 9
            MsgBox "Error: " & Err.Number & ", " & _
            Err.Description & vbCrLf & "Finaliza...", vbCritical, "Error"
            evaluarError = Finalizar
            
        'disco no preparado
        Case 71
            res = MsgBox("Error: " & Err.Number & ", " & _
            Err.Description, vbCritical + vbRetryCancel, "Error")
            If res = vbRetry Then
                evaluarError = Reintentar
            Else: evaluarError = Cancelar
            End If
        
    End Select
End Function


