Attribute VB_Name = "G_DateFunctions"
Option Explicit
Option Compare Database

'-----------------------------------------------------------------------------
'   Name:   FechaDelLunes
'   Retr:   Devuelve la fecha del primer lunes
'   Desc:   Calcula la fecha del primer lunes de la semana en que nos
'           encontramos.
'-----------------------------------------------------------------------------
Public Function FechaDelLunes() As Date
    Dim lunes As Date
    
    lunes = Date
    
    While WeekdayName(DatePart("w", lunes, vbMonday), , vbMonday) <> "lunes"
        lunes = DateAdd("d", -1, lunes)
    Wend
    FechaDelLunes = lunes
End Function
