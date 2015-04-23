Attribute VB_Name = "SIFOC_Tarea"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/06/2011 - Actualización:  15/06/2011
'   Name:   insTarea
'   Desc:   Inserta tarea en la tabla
'   Param:  idPersona(long), i
'
'   Retur:  String con apellido y nombre de la persona activa
'---------------------------------------------------------------------------
Public Function insTarea(idPersona As Long, _
                         idIfocUsuario As Long, _
                         tarea As String, _
                         Optional descripcion As String = "", _
                         Optional idTareaPrioridad As Integer = 0, _
                         Optional fechaLimite As Date = "01/01/1900", _
                         Optional realizado As Integer = 0, _
                         Optional updDate As Date = "01/01/1900")
    Dim strSql As String
    Dim strFields As String
    Dim strValues As String
On Error GoTo TratarError
    
    strFields = "(fkPersona" & _
                ", fkIfocUsuario" & _
                ", tarea" & _
                IIf(descripcion = "", "", ", descripcion") & _
                IIf(idTareaPrioridad = 0, "", ", fkTareaPrioridad") & _
                IIf(fechaLimite = 0, "", ", fkCurso") & _
                IIf(realizado = 0, "", ", fkCursoNivel") & _
                IIf(updDate = "", "", ", updDate") & _
                ")"
    
    strValues = "(" & idPersona & _
                ", " & idIfocUsuario & _
                ", '" & filterSQL(tarea) & "'" & _
                IIf(descripcion = "", "", ", '" & filterSQL(descripcion) & "'") & _
                IIf(idTareaPrioridad = 0, "", ", " & idTareaPrioridad) & _
                IIf(fechaLimite = "", "", "#" & Format(fechaLimite, "mm/dd/yyyy hh:nn:ss") & "#") & _
                IIf(realizado = 0, "", ", " & realizado) & _
                IIf(updDate = "", "", ", #" & updDate & "#") & _
                ")"
    
    strSql = " INSERT INTO t_tarea" & _
             strFields & _
             " VALUES " & _
             strValues & ";"

TratarError:

End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/06/2011 - Actualización:  15/06/2011
'   Name:   openFrmGestionTarea_persona
'   Desc:   Abre formulario GestionTarea
'   Param:  idPersona(long), identificador de persona
'
'   Retur:  String con apellido y nombre de la persona activa
'---------------------------------------------------------------------------
Public Function openFrmGestionTarea_persona(idPersona As Long)
    Dim db As dao.database
    Dim rst As dao.Recordset
    Dim strSql As String
    
    Dim frm As New Form_GestionTarea
    
    If (idPersona <> 0) Then
        frm.setIdPersona (idPersona)
        frm.actualizaForm
    End If
    
    'frm.NavigationButtons = False
    frm.visible = True
    
    LigaFormulario.LigaFormulario frm
    
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/06/2011 - Actualización:  15/06/2011
'   Name:   openFrmGestionTareasUsuario_persona
'   Desc:   Abre formulario GestionTareasUsuario
'   Param:  idPersona(long), identificador de persona
'
'   Retur:  String con apellido y nombre de la persona activa
'---------------------------------------------------------------------------
Public Function openFrmGestionTareasUsuario_persona(idPersona As Long)

    Dim db As dao.database
    Dim rst As dao.Recordset
    Dim strSql As String
    
    Dim frm As New Form_GestionTareasUsuario
    
    frm.setIdPersona (idPersona)
    frm.actualizaForm
    
    'frm.NavigationButtons = False
    frm.AllowAdditions = True
    frm.AllowEdits = True
    frm.visible = True
    
    LigaFormulario.LigaFormulario frm
    
End Function
