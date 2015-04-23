Attribute VB_Name = "z_test"
Option Explicit
Option Compare Database

Public Function DemandaEliminaDuplicados()
    Dim rsDemDup As ADODB.Recordset
    Dim rsDemTabla As ADODB.Recordset
    Dim str1 As String
    Dim str2 As String
    
    str1 = " SELECT t_cnodebusqueda.fkPersona, t_cnodebusqueda.fkCno2011, t_cnodebusqueda.nivel, Count(t_cnodebusqueda.fkCno) AS NumCnoDuplicados" & _
           " FROM t_cnodebusqueda" & _
           " GROUP BY t_cnodebusqueda.fkPersona, t_cnodebusqueda.fkCno2011, t_cnodebusqueda.nivel" & _
           " HAVING (((Count(t_cnodebusqueda.fkCno))>1));"
        
    Set rsDemDup = New ADODB.Recordset
    rsDemDup.Open str1, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    Set rsDemTabla = New ADODB.Recordset
    
    Dim idPersona As Long
    Dim idCno As Long
    Dim nivel As String
    Dim numDel As Integer
    Dim count As Integer
    count = 0
    If Not rsDemDup.EOF Then
        rsDemDup.MoveFirst
        'Query duplicate demands
        While Not rsDemDup.EOF
            idPersona = rsDemDup!fkPersona
            idCno = rsDemDup!fkCno2011
            nivel = rsDemDup!nivel
            numDel = rsDemDup!numCnoDuplicados - 1
            str2 = " SELECT fkPersona, fkCno2011, nivel" & _
                   " FROM t_cnodebusqueda" & _
                   " WHERE fkPersona =" & idPersona & " AND fkCno2011=" & idCno & " AND nivel =" & nivel & _
                   " ORDER BY nivel DESC"
            rsDemTabla.Open str2, CurrentProject.Connection, adOpenDynamic, adLockOptimistic
            If Not rsDemTabla.EOF Then
                rsDemTabla.MoveFirst
                'Table delete duplicate records of demand
                While Not rsDemTabla.EOF
                    If numDel > 0 Then
                        rsDemTabla.delete adAffectCurrent
                        numDel = numDel - 1
                        count = count + 1
                    End If
                    rsDemTabla.MoveNext
                Wend
            End If
            rsDemTabla.Close
            rsDemDup.MoveNext
        Wend
    End If
    
    rsDemDup.Close
    Set rsDemDup = Nothing
    Set rsDemTabla = Nothing
    Debug.Print "Regs eliminados: " & count
End Function

Public Function h(sec As Integer)
    Dim now As Long
      now = Timer()
      MsgBox "espera " & sec & " para continuar"
      Do
          DoEvents
      Loop While (Timer < now + sec)
      Debug.Print "out"
End Function

Public Function updateServicioIntermediacionEmpresas()
    Dim rs As ADODB.Recordset
    Dim str As String
    Dim fecha As Date
    Dim newFecha As Date
    Dim iniFecha As Date
    Dim finFecha As Date
    Dim idOrganizacion As Long
    Dim idOk As String
    Dim idKo As String
    Dim idOld As String
    Dim counter As Integer
    Dim OBS As String
    
    str = "___comor1"
    
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        rs.MoveFirst
    End If
    
    OBS = "Alta automática masiva, cogemos fecha 1ª oferta del año como fecha inicio y 31/12/año como fecha fin, para actualizar alta servicio Intermediacion de empresa."
    
    counter = 0
    While Not rs.EOF
        idOrganizacion = rs!idOrganizacion
        fecha = rs!minfecha
        newFecha = DateSerial(Year(fecha), 12, 15)
        iniFecha = DateSerial(Year(fecha), 1, 1)
        finFecha = DateSerial(Year(fecha), 12, 31)
        If Not isUserServiceActiveInInterval(iniFecha, finFecha, 29, , idOrganizacion) Then
            If altaServicioOrganizacion(idOrganizacion, 29, fecha, 14, finFecha, 19, , OBS) = 0 Then
                idOk = idOk & ", " & idOrganizacion & "(" & newFecha & ")"
                counter = counter + 1
            Else
                idKo = idKo & ", " & idOrganizacion & "(" & newFecha & ")"
            End If
        Else
            idOld = idOld & ", " & idOrganizacion & "(" & newFecha & ")"
        End If
        rs.MoveNext
    Wend
    
    Debug.Print "Contador:" & counter
    Debug.Print "idOk:" & idOk
    Debug.Print "idKo:" & idKo
    Debug.Print "idOld:" & idOld
    
    rs.Close
    Set rs = Nothing
End Function
