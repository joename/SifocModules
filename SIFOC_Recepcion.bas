Attribute VB_Name = "SIFOC_Recepcion"
Option Explicit
Option Compare Database

Public Static Function getInfoRecepcion(idIfocUsuario As Long, fechaInicio As Date, fechaFin As Date) As Variant
    Dim num As Integer
    Dim strSql As String
    Dim rs As dao.Recordset
    Dim fechaI, FECHAF As Date
    Dim arrRecepcion(3) As Integer
    
    fechaI = Format(fechaInicio, "mm/dd/yyyy")
    FECHAF = Format(fechaFin, "mm/dd/yyyy") & " 23:59:59"

    strSql = " SELECT Count(t_recepcion.id) as altas" & _
             " FROM t_recepcion" & _
             " WHERE (fechaHora <= #" & FECHAF & "# AND fechaHora >= #" & fechaI & "#)" & _
             " AND (fkIFOCUsuario =" & idIfocUsuario & ")"
    
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'Set rs = New ADODB.Recordset
    'rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        'Acciones realizadas, Tiempo
        arrRecepcion(1) = Nz(rs!altas, 0)
    End If
        
    rs.Close
    Set rs = Nothing
    
    getInfoRecepcion = arrRecepcion
    
End Function

