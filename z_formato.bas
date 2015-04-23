Attribute VB_Name = "z_formato"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  21/1/2009 - Actualización:  21/1/2009
'   Name:   dni
'   Desc:   Calcula el ultimo digito(letra) del DNI o NIE
'   Param:  Dni_Nie, string ( 8 numeros)
'                 formato dni: NNNNNNNN
'                         nie: NNNNNNNN
'   Retur:  devuelve dniNie correcto NNNNNNNL
'---------------------------------------------------------------------------


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
Private Sub pruebaTablaTMP()

Dim cn As ADODB.Connection
Dim cmd As ADODB.command
Dim rst As ADODB.Recordset

'Inicializar conexión
Set cn = New ADODB.Connection
cn.Open getMyConnectionString
Set rst = New ADODB.Recordset
Set cmd = New ADODB.command


With cmd
    'Inicializar propiedades del comando
    .ActiveConnection = cn
    .CommandText = "memorias_BaseInserciones ('2007-1-1','2007-12-31')" 'se pueden pasar así directamente o con el metodo .CreateParameter (ver abajo)
    .CommandType = adCmdStoredProc
     
    ' codigo para input variables
    '.Parameters.Append .CreateParameter("prmFecIni", adDate, adParamInput, , "2007-1-1")
    '.Parameters.Append .CreateParameter("prmFecIni", adDate, adParamInput, , "2007-12-31")
    '.Parameters.Append .CreateParameter("dob", adInteger, adParamInput, 11, 2)
    ' output variables
    '.Parameters.Append .CreateParameter("prmOut", adVarChar, adParamOutput, 45)
    '.Parameters.Append .CreateParameter("prmOut", adVarChar, adParamOutput, 45)
    
    'Ejecutar el procedimiento almacenado
    Set rst = .Execute

 
    'codigo para extraer la  variable output
    'MsgBox .Parameters("prmOut")
    'MsgBox rst.fields.count & ", " & rst.fields.Item(0)

End With

   'cerrar objetos
   rst.Close
   Set rst = Nothing
   
End Sub

