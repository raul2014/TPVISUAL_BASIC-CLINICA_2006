Attribute VB_Name = "mdlValCons"
Option Explicit
Public rtaError As AccionError
Public conn As ADODB.Connection
Public Const ALTA = 1
Public Const MODIFICACION = 2

'Esta variable contendra el codigo de usuario que esta usando la aplicacion
' ademas con este dato podre dejar sentado que usuario modifica o da de alta los datos de los formularios
Public usuarioActual As Integer


