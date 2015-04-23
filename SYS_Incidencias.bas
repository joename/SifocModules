Attribute VB_Name = "SYS_Incidencias"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  16/07/2009
'   Name:
'   Desc:   Crea una incidencia nueva de SIFOC
'   Param:  Dni_Nie, string ( 8 numeros)
'                 formato dni: NNNNNNNN
'                         nie: NNNNNNNN
'   Retur:  devuelve dniNie correcto NNNNNNNL
'---------------------------------------------------------------------------
Public Function addIncidencia(formulario As String, _
                              tipo As String, _
                              openargs As String, _
                              incidencia As String) As Integer
On Error GoTo TratarError
    'Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim cnStr As String
    Dim str As String
    Dim fecha As String
    Dim myTipo As String
    Dim myIncidencia As String
    
    Dim cmd As ADODB.command
    
    'cnStr = connectionStr(, _
                          SERVER, _
                          , _
                          DDB1, _
                          USER_DB1, _
                          PASS_DB1)
    'cn.Open cnStr
    
    fecha = Format(now(), "mm/dd/yyyy hh:nn:ss")
    myIncidencia = filterSQL(incidencia)
    
    str = " INSERT INTO sysincidencias (fecha, form, openargs, tipo, incidencia, fkIfocUsuario, pc)" & _
          " VALUES (now(), '" & formulario & "', '" & openargs & "','" & tipo & _
                    "', """ & myIncidencia & """, " & U_idIfocUsuarioActivo & ", '" & computerName() & " - " & userName() & "')"
Debug.Print str
    Set cmd = New ADODB.command
    
    With cmd
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdText
        .CommandText = str
        .Execute
    End With
    
    Set cmd = Nothing
    
    addIncidencia = 0
    
    Exit Function
TratarError:
    addIncidencia = -1
    MsgBox "Error: " & Err.description, vbOKOnly, "Alert: " & formulario
    debugando Err.description
End Function
