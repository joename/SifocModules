Attribute VB_Name = "G_StoredProcedure"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  José Espases    AutorAct: -
'   Fecha:  14/02/2011  FechaAct: -
'   Name:   procedimientoAlmacenado
'   Desc:   Ejecuta un procedimiento almacenado en MySQL y devuelve datos en forma de tablas ACCESS.
'           las tablas ACCESS deberán borrarse después utilizando el procedimiento deleteTable(tableName)
'   Param:  strStoredProcedure AS String, contiene el nombre del procedimiento almacenado
'           strParametros As String, contiene los parámetros separados por comas
'           tableNames As String, contiene el número de tablas a devolver y sus nombres separados por ##
'   Ejemplo: procedimientoAlmacenado("procName", "23, 'texto'", "2##tableName1##tableName2")
'           En este ejemplo los parámetros son un número y un texto, y devuelve 2 tablas
'   Retur:  string, cadena de caracteres generado aleatoriamente
'---------------------------------------------------------------------------
Public Function procedimientoAlmacenado(strStoredProcedure As String, _
                                        strParametros As String, _
                                        tableNames As String) As Boolean
    Dim cn As ADODB.Connection
    Dim cmd As ADODB.command
    Dim x As Integer
    Dim tableName As String
    Dim rstName As Variant
    Dim rst As ADODB.Recordset
    
    'Debug.Print strParametros
    procedimientoAlmacenado = True
    rstName = Split(tableNames, "##")
    
    Set cn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    Set cmd = New ADODB.command
    
    Debug.Print getMyConnectionString()
    
    cn.Open getMyConnectionString()
    cmd.ActiveConnection = cn
    cmd.CommandText = strStoredProcedure & " (" & strParametros & ")"
    cmd.CommandType = adCmdStoredProc

    Set rst = cmd.Execute
    For x = 1 To rstName(0)
        'Crea la Tabla a partir del RecordSet generado en el Stored Procedure
        tableName = rstName(x)
        procedimientoAlmacenado = makeTable(tableName, rst)
        ' move to next resultset, using nextRecordSet()
        If x < rstName(0) Then
            Set rst = rst.NextRecordset
        End If
    Next
    
    'rst.Close
    Set rst = Nothing
    cn.Close
    Set cn = Nothing

End Function

'-------------------------------------------------------------------------------------------------------------------------------------
'CREATE DEFINER=`root`@`localhost` PROCEDURE `Memorias_BaseInserciones`(IN prmFecIni datetime, IN prmFecFin datetime)
'BEGIN
'        SELECT Inserciones.fkPersona,
'                 Count(Inserciones.id) AS Inserciones,
'                 Inserciones.sexo,
'                 Inserciones.fechaNacimiento,
'                 Max(Inserciones.fechaInicio) As InicioUltimaInsercion
'                 FROM (SELECT t_insercion.fkPersona,
'                          a_sexo.sexo,
'                          t_persona.fechaNacimiento,
'                          t_insercion.fechaInicio,
'                          t_insercion.fechaFin,
'                          t_insercion.id
'                        from t_insercion
'                        INNER JOIN (t_persona INNER JOIN a_sexo ON t_persona.fkSexo = a_sexo.id)
'                                 ON t_insercion.fkPersona = t_persona.id
'                        WHERE (((t_insercion.fechaInicio)<=prmFecFin)
'                                AND ((t_insercion.fechaFin)>=prmFecIni
'                                Or (t_insercion.fechaFin) Is Null))) as Inserciones
'        GROUP BY Inserciones.fkPersona, Inserciones.sexo, Inserciones.fechaNacimiento;
'End
'-------------------------------------------------------------------------------------------------------------------------------------
'Desde Access
'-------------------------------------------------------------------------------------------------------------------------------------
'Cadena de conexion:
'Const DB_CONNECT As String = "Driver=SIFOCLocal;Server=localhost;Port=3306;Database=sifoc;User=root; Password=root;Option=3;"

'-----------------------------------------------------------------------------------------
'   Name:   btn_ProcedimientoAlmacenado_Click
'   Return:
'   Obs:    ejecuta un procedimiento almacenado de MySQL obteniendo el recordset resultante.
'   Comentarios: el recordset no puede asignarse a un formulario.
'               por ejemplo la instrucción:    set frm.recordset= rs
'               donde frm es un formulario y rs es el recordset obtenido del proc. almacenado
'               da un error en el ODBC
'-----------------------------------------------------------------------------------------
