Public Module literals
    'General usage
    Public Const SAVEAS As String = "Salvar como ..."

    'User Messages
    Public Const ERRORLOADINGDATA As String = "Error al cargar datos: "
    Public Const ERRORSAVE As String = "Hubo un error al salvar el registro. Favor reintente. Puede obtener más detalles en la bitácora de errores"
    Public Const ERRORNUMERICONLY As String = "El campo solo acepta valores numéricos"
    Public Const ERRORDATANOTGENERATED As String = "Los datos no fueron generados"
    Public Const ERRORDUPLICATERECORD As String = "El registro se encuentra duplicado"
    Public Const MSGSUCCESSSAVE As String = "El registro ha sido salvado"
    Public Const MSGPRINTRECORDS As String = "¿Confirma que se imprimirá(n) {0} registro(s)?"
    Public Const ERRORPRINTING As String = "No se logro imprimir el registro, favor intente nuevamente"
    Public Const ERROR_NUMERIC_ONLY As String = "El valor debe ser numérico"

    'Database related errors
    Public Const ERRROOLEDBTYPENOTFOUND As String = "El System.Type {0} no tiene un tipo OleDB compatible en el campo {1}"
    Public Const ERRROMYSQLTYPENOTFOUND As String = "El System.Type {0} no tiene un tipo MySQL compatible en el campo {1}"
    Public Const ERROR_NO_ID_FIELDS As String = "No se ha definido la llave principal para la tabla: "
    Public Const ERROR_CONNECT As String = "Error al abrir conexión a base de datos: "
    Public Const ERROR_GET_DATA As String = "Error al recuperar registros remotos"
    Public Const ERROR_GET_LOCAL_STRUCTURE As String = "Error al recuperar estructura de tabla"
    Public Const ERROR_UPDATE_STATUS As String = "Error actualizando estado de "
    Public Const ERROR_UPDATE As String = "Error al actualizar "
    Public Const ERROR_LOAD_RECORD As String = "Error al descargar registro"
    Public Const ERROR_LOAD_UPDATE As String = "Error al descargar actualización de registro"
    Public Const ERROR_CONVERSION As String = "Error al convertir datos"
    Public Const ERROR_EXECUTING As String = "Error al ejecutar comando"
    Public Const ERROR_SELECT As String = "Error al seleccionar registros"
    Public Const ERROR_TYPE_NOT_DEFINED As String = "El tipo de datos {0} no está definido"

    'Database related messages
    Public Const MSG_GET_RECORDS As String = "Recuperando registros: "
    Public Const MSG_INSERT_RECORD As String = "Cargando registro: "
    Public Const MSG_UPDATE_RECORD As String = "Actualizando registro: "
    Public Const MSG_CONNECT_DB As String = "Abriendo conexion a base de datos"
    Public Const MSG_GET_DATA As String = "Recuperando registros remotos"
    Public Const MSG_GET_STATUS As String = "Recuperando estado de "
    Public Const MSG_GET_STRUCTURE As String = "Recuperando estructura de "
    Public Const MSG_UPDATE_STATUS As String = "Actualizando estado de "
    Public Const MSG_PROCESS_RECORD As String = "Procesando registro "
    Public Const MSG_GET_INDEX As String = "Obteniendo indice de sincronización"
    Public Const MSG_GET_FIELDS As String = "Construyendo listado de campos"
    Public Const MSG_GET_PARAMETERS As String = "Construyendo listado de parametros"
    Public Const MSG_GET_CONDITIONS As String = "Construyendo listado de condiciones"
    Public Const MSG_GET_COMMAND As String = "Construyendo comando de listado de condiciones"

    'SFTP Related Messaged
    Public Const ERRORUPLADINGFILE As String = "Error al cargar archivo: {0}"
    Public Const ERRORDOWNLOADINGFILE As String = "Error al descargar archivo: {0}"
    Public Const ERRORGETDIRECTORYLIST As String = "Error al recuperar listado de directorio {0}"

    'SQL Server Errors
    Public Const ERROR_SQL_CLIENT_NOT_INSTALLED As String = "No se encuentra instalado el cliente de SQL Server"
    Public Const ERROR_SQL_CLIENT_INITIALIZATION As String = "No se logró iniciar el cliente de conexión. Si está utilizando SQL Server asegurese de tener instalado el Cliente de SQL: "
    Public Const ERROR_DATABASE_PARAMETER_NOT_SET As String = "Los parametros de conexion no han sido configurados: "

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

    'Transfer Data
    Public Const ERRORCONNECTDESTINYSERVER As String = "Error al conectar al servidor destino: "
    Public Const ERRORCONNECTSOURCESERVER As String = "Error al conectar al servidor destino: "
    Public Const ERRORDIFFERENTSCHEMAS As String = "El esquema de la tabla fuente no es el mismo que el de la tabla destino"

    'Windows Services
    Public Const ERRORSERVICENOTFOUND As String = "Servicio no encontrado; asegurese de que este instalado: {0}"
    Public Const ERRORSTARTINGSERVICE As String = "Error al iniciar el servicio: {0}"

    'HTTP Server
    Public Const HTTPSYSTEMREQUIREMENTS As String = "Windows XP SP2, Server 2003, o superior es requerido para utilizar HTTP Listener."
    Public Const HTTPMETHODNOTFOUND As String = "El metodo no fue encontrado"

    'GDI
    Public Const GDIUNKNOWNFORMAT As String = "Formato desconocido: "

    'Dates
    Public Const JAN As String = "Enero"
    Public Const FEB As String = "Febrero"
    Public Const MAR As String = "Marzo"
    Public Const APR As String = "Abril"
    Public Const MAY As String = "Mayo"
    Public Const JUN As String = "Junio"
    Public Const JUL As String = "Julio"
    Public Const AUG As String = "Agosto"
    Public Const SEP As String = "Septiembre"
    Public Const OCT As String = "Octubre"
    Public Const NOV As String = "Noviembre"
    Public Const DEC As String = "Diciembre"

    'User Errors
    Public Const ERROR_FORMAT_TELEPHONE As String = "El teléfono no está en el formato correcto"
    Public Const ERROR_FORMAT_EMAIL As String = "El correo electrónico no está en el formato correcto"
    Public Const ERROR_PERMISION_AREA_MANAGER As String = "Solo un gerente de área puede habiltar este permiso"
    Public Const ERROR_PERMISION_ENTRY As String = "Su usuario no dispone de los permisos necesarios para entrar al aplicativo"
    Public Const ERROR_PERMISION_OPTION As String = "Su usuario no dispone de los permisos necesarios para esta opción"
    Public Const ERROR_PERMISION_BRANCH As String = "Esta sucursal no puede ejecutar esta opción"
    Public Const ERROR_SELECT_RECORD As String = "Debe seleccionar al menos un registro para utilizar esta opción"
    Public Const ERROR_INVALID_INPUT As String = "Debe ingresar un valor válido"

    'Configuration Constants
    Public Const MAILESCAPE As String = "@." 'Represents no mail given

    'Process Handling
    Public Const WARN_MAX_RUNS As String = "Se alcanzó la máxima cantidad de ciclos para un proceso: "
    Public Const MSG_CONNECT_REMOTE_DB As String = "Conectando a base de datos remota "
    Public Const ERROR_CONNECT_REMOTE_DB As String = "Error al conectar a base de datos remota"

    'File transfers
    Public Const MSG_DOWNLOADING_FILE As String = "Error descargando archivo"
    Public Const MSG_DELETING_FILE As String = "Eliminando archivo"
    Public Const ERROR_DELETING_FILE As String = "No se pudo eliminar el archivo {0}. Asegurese que este accesible para borrarse"

    'File Related Errors
    Public Const ERROR_PROCESSING_FILE As String = "Error al procesar archivo"

    'Generic Fields
    Public Const GF_NAME As String = "Nombre"
    Public Const GF_CUSTOMER As String = "Cliente"
    Public Const GF_CUSTOMERS As String = "Clientes"
    Public Const GF_CODE As String = "Código"
    Public Const GF_DESCRIPTION As String = "Descripción"
    Public Const GF_PRODUCT_CODE As String = "Código de Producto"
    Public Const GF_PRODUCT As String = "Producto"
    Public Const GF_SUPPLIER As String = "Proveedor"
    Public Const GF_DATE As String = "Fecha"
    Public Const GF_QUANTITY As String = "Cantidad"
    Public Const GF_AMOUNT As String = "Monto"
    Public Const GF_PRICE As String = "Precio"
    Public Const GF_EXISTANCE As String = "Existencia"
    Public Const GF_BARCODE As String = "Código de Barras"
    Public Const GF_SALES_TAX As String = "ISV"
    Public Const GF_BRANCH As String = "Sucursal"
    Public Const GF_INVOICE As String = "Factura"
    Public Const GF_USER As String = "Usuario"
    Public Const GF_EMAIL As String = "E-Mail"
    Public Const GF_TAX As String = "Impuesto"
    Public Const GF_STATUS As String = "Estado"
    Public Const GF_LENGTH As String = "Longitud"
    Public Const GF_STOCK As String = "Inventario"
    Public Const GF_VALUE As String = "Valor"
    Public Const GF_NUMBER As String = "Número"
    Public Const GF_IS_ACTIVE As String = "Activo"
End Module
