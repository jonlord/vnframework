﻿Public Module literals
    'User Messages
    Public Const ERRORLOADINGDATA As String = "Error al cargar datos: "
    Public Const ERRORSAVE As String = "Hubo un error al salvar el registro. Favor reintente. Puede obtener más detalles en la bitácora de errores"
    Public Const ERRORNUMERICONLY As String = "El campo solo acepta valores numéricos"
    Public Const ERRORDATANOTGENERATED As String = "Los datos no fueron generados"
    Public Const ERRORDUPLICATERECORD As String = "El registro se encuentra duplicado"
    Public Const MSGSUCCESSSAVE As String = "El registro ha sido salvado"

    'Database related messages
    Public Const ERRROOLEDBTYPENOTFOUND As String = "El System.Type {0} no tiene un tipo OleDB compatible en el campo {1}"
    Public Const ERRROMYSQLTYPENOTFOUND As String = "El System.Type {0} no tiene un tipo MySQL compatible en el campo {1}"

    'SFTP Related Messaged
    Public Const ERRORUPLADINGFILE As String = "Error al cargar archivo: {0}"
    Public Const ERRORDOWNLOADINGFILE As String = "Error al descargar archivo: {0}"
    Public Const ERRORGETDIRECTORYLIST As String = "Error al recuperar listado de directorio {0}"

    'SQL Server Error Messaged replacements
    Public Const REPLACENULLFIELD As String = "No puede dejar el campo '$1' vacio"
    Public Const REPLACEUNCLOSEDQUOTATION As String = "El valor ingresado tiene caracteres invalidos"
    Public Const REPLACEARITHMETICOVERFLOW As String = "Al menos uno de los valores ingresados es mayor al permitido."
    Public Const REPLACESTRINGTRUNCATED As String = "Ha ingresado un valor con una longitud mayor a la permitida"
    Public Const REPLACEFOREIGNKEY As String = "The ALTER TABLE statement conflicted With the FOREIGN KEY constraint"
    Public Const REPLACEMISSINGFIELD As String = "Su base de datos puede estar desactualizada. Favor proceda a actualizarla para obtener el campo "
    Public Const REPLACEMISSINGTABLE As String = "Su base de datos puede estar desactualizada. Favor proceda a actualizarla para obtener la tabla "
    Public Const REPLACEFIELDCOUNT0 As String = "Su base de datos puede estar desactualizada. Favor proceda a actualizarla para obtener los campos de la tabla "
    Public Const REPLACEFOREIGNKEYINSERT As String = "No se encuentran todos los registros necesarios en la tabla $4"
End Module
