Attribute VB_Name = "G_Centralita"
Option Explicit
Option Compare Database

Const TIEMPOREBOTE As Integer = 12
Const TIEMPOPERDIDA As Integer = 60 'Segundos, se considera perdida definitiva ¿?

'--------------------------------------------------------------------------------
'   Name: CargaLlamadasINTablaLocal
'   Desc: Copia las llamadas entrantes en el periodo determinado
'         a la tabla LOCAL L_calldata
'   Parm: fechaI,   fecha de inicio
'
'   Retr: Devuelve el número de perdidas eliminado, 0 si no elimina ninguno.
'--------------------------------------------------------------------------------
Public Function cargaLlamadasINTablaLocal(fechaInicio As Date, _
                                          fechaFin As Date) As Integer
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim cnStr As String
    Dim str As String
    Dim str1 As String
    Dim fechaI As String
    Dim FECHAF As String
    
    cnStr = getConnectionString(, _
                                "serverifoc", _
                                , _
                                "utilerep", _
                                "char", _
                                "userchar")
    cn.Open cnStr
    
    fechaI = Format(fechaInicio, "yyyy/mm/dd hh:nn:ss")
    FECHAF = Format(fechaFin, "yyyy/mm/dd hh:nn:ss")
    
    str = " SELECT CALLID, CALLTYPE, CALLDATE, CALLED, CALLER, CALLDUR, CALLRESP" & _
          " FROM calldata" & _
          " WHERE (((calldata.CALLTYPE)=2) AND ((calldata.CALLDATE) Between '" & fechaI & "' And '" & FECHAF & " 23:59:59'))" & _
          " ORDER BY calldata.CALLDATE DESC;"
    
    rs.Open str, cn
    
    'Borramos tabla local
    CurrentDb.Execute "DELETE * FROM L_calldata;"
    
    str1 = "L_calldata"
    
    Set rs1 = New ADODB.Recordset
    rs1.Open str1, CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTable
    
    If Not rs.EOF Then
        rs.MoveFirst
    End If
    
    While Not rs.EOF
        rs1.AddNew
        
        rs1!CALLID = rs!CALLID
        rs1!CALLTYPE = rs!CALLTYPE
        rs1!CALLDATE = rs!CALLDATE
        rs1!called = rs!called
        rs1!CALLER = rs!CALLER
        rs1!CALLDUR = rs!CALLDUR
        rs1!CALLRESP = rs!CALLRESP
        rs1.update
        
        rs.MoveNext
    Wend
    
    'Cerramos recordsets
    rs1.Close
    Set rs1 = Nothing
    
    rs.Close
    Set rs = Nothing
End Function

'--------------------------------------------------------------------------------
'   Name: EliminaRebotes
'   Desc: Elimina las llamadas que rebotan a otra extensión de la 212,221 y 310
'         Secuéncia de rebote 212 -> 221 -> 310
'         Rebote se realiza tras 12 seg (3 tonos)
'         Si el número comunica pasa al siguiente y no cuenta como perdida
'   Parm: fechaI,   fecha de inicio
'
'   Retr: Devuelve el número de perdidas eliminado, 0 si no elimina ninguno.
'--------------------------------------------------------------------------------
Public Function EliminaRebotes(fechaInicio As Date, _
                               fechaFin As Date) As Integer
    Dim rs As ADODB.Recordset
    Dim str As String
    Dim contador As Integer 'número de posibles rebotes
    Dim callid1 As Double 'id llamada anterior
    Dim callDate1 As Date  'date llamada anterior
    Dim posibleRebote As Boolean
    
    'CALLTYPE = 2 (Llamadas entrantes)
    'CALLED   = teléfono/extensión que recibe la llamada
    'CALLDATE = fecha y hora de la llamada
    'CALLER   = teléfono/extensión que realiza la llamada
    'CALLDUR  = duración de la llamada (seg.)
    'CALLRESP = tiempo antes de coger llamada (seg.)
    str = " SELECT CALLID, CALLTYPE, CALLDATE, CALLED, CALLER, CALLDUR, CALLRESP" & _
          " FROM L_calldata" & _
          " WHERE (CALLTYPE=2)" & _
          " AND ((CALLDATE) Between #" & Format(fechaInicio, "mm/dd/yyyy") & "# And #" & Format(fechaFin, "mm/dd/yyyy") & " 23:59:59#)" & _
          " AND ((CALLED)='212' Or (CALLED)='221' Or (CALLED)='220'Or (CALLED)='309'Or (CALLED)='310')" & _
          " ORDER BY CALLDATE ASC;"

'debugando str
    
    'Abrimos el recordset
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        rs.MoveFirst
        contador = 0
        posibleRebote = False
        Do While Not rs.EOF
            If (posibleRebote) Then
                'Eliminamos llamadas que
                If (DateDiff("s", rs!CALLDATE, callDate1) <= TIEMPOREBOTE) _
                    And (rs!called = 212 Or _
                         rs!called = 221 Or _
                         rs!called = 220 Or _
                         rs!called = 309 Or _
                         rs!called = 310) Then
                    str = " DELETE *" & _
                          " FROM L_calldata" & _
                          " WHERE CALLID= " & callid1 & ";"
                    CurrentDb.Execute str
                    contador = contador + 1
                End If
                posibleRebote = False
            End If
            'Miramos si es un posible rebote
            If rs!CALLDUR = 0 And rs!CALLRESP <= 12 _
               And (rs!called = 212 Or _
                    rs!called = 221 Or _
                    rs!called = 220 Or _
                    rs!called = 309 Or _
                    rs!called = 310) Then
                posibleRebote = True
                callid1 = rs!CALLID
                callDate1 = rs!CALLDATE
            End If
            'For Each rs In rs
            'Next rs.MoveNext
            rs.MoveNext
        Loop
    End If
    
    'Cerramos el recordset
    rs.Close
    Set rs = Nothing
    
    'Devolvemos el número de llamadas perdidas eliminadas
    EliminaRebotes = contador
End Function

'--------------------------------------------------------------------------------
'   Name: EliminaLlamadasNOPerdidas
'   Desc: Elimina las llamadas que no superan el TIEMPOPERDIDA establecido
'         por el IFOC para que no contabilice como perdida
'         Rebote se realiza tras 12 seg (3 tonos)
'   Parm: fechaI,   fecha de inicio
'         fechaF,   fecha de fin
'   Retr: Devuelve el número de ¿perdidas? eliminado, 0 si no elimina ninguno.
'--------------------------------------------------------------------------------
Public Function EliminaLlamadasNOPerdidas(fechaInicio As Date, _
                                          fechaFin As Date) As Integer
    Dim rs As ADODB.Recordset
    Dim str As String
    Dim contador As Integer 'número de posibles rebotes
    Dim tiempo As Integer
    
    'CALLTYPE = 2 (Llamadas entrantes)
    'CALLED   = teléfono/extensión que recibe la llamada
    'CALLDATE = fecha y hora de la llamada
    'CALLER   = teléfono/extensión que realiza la llamada
    'CALLDUR  = duración de la llamada (seg.)
    'CALLRESP = tiempo antes de coger llamada (seg.)
    str = " SELECT CALLID, CALLTYPE, CALLDATE, CALLED, CALLER, CALLDUR, CALLRESP" & _
          " FROM L_calldata" & _
          " WHERE (CALLTYPE=2)" & _
          " AND ((CALLDATE) Between #" & Format(fechaInicio, "mm/dd/yyyy") & "# And #" & Format(fechaFin, "mm/dd/yyyy") & " 23:59:59#)" & _
          " AND ((CALLED)='212' Or (CALLED)='221' Or (CALLED)='220'Or (CALLED)='309'Or (CALLED)='310')" & _
          " AND (CALLDUR = 0)" & _
          " ORDER BY CALLDATE ASC;"

'debugando str
    
    'Abrimos el recordset
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        rs.MoveFirst
        contador = 0
        Do While Not rs.EOF
            If rs!called = "212" Then
                tiempo = rs!CALLRESP
            ElseIf rs!called = "221" Then
                tiempo = rs!CALLRESP + 12
            ElseIf rs!called = "220" Then
                tiempo = rs!CALLRESP + 24
            ElseIf rs!called = "309" Then
                tiempo = rs!CALLRESP + 36
            ElseIf rs!called = "310" Then
                tiempo = rs!CALLRESP + 48
            End If
            
            If tiempo < 60 Then
                contador = contador + 1
                
                str = " DELETE * FROM L_CALLDATA WHERE CALLID=" & rs!CALLID & ";"
                CurrentDb.Execute str
            End If
            
            rs.MoveNext
        Loop
    End If
    
    'Cerramos el recordset
    rs.Close
    Set rs = Nothing
    
    'Devolvemos el número de llamadas perdidas eliminadas
    EliminaLlamadasNOPerdidas = contador
End Function
