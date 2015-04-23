Attribute VB_Name = "SIFOC_Imagen"
Option Explicit
Option Compare Database

Public Const rutaImagen As String = "\\serverifoc\Archivos_SIFOC$\PersonasFoto\"

Public Function EscogeImagen() As String
  '© Juan M. Afán de Ribera
  Dim wzhwndOwner As Long
  Dim wzAppName As String
  Dim wzDlgTitle As String
  Dim wzOpenTitle As String
  Dim wzFile As String
  Dim wzInitialDir As String
  Dim wzFilter As String
  Dim wzFilterIndex As Long
  Dim wzView As Long
  Dim wzflags As Long
  Dim wzfOpen As Boolean
  Dim ret As Long
    
    WizHook.Key = 51488399
    
    wzhwndOwner = 0&
    wzAppName = ""
    wzDlgTitle = "Escoge la imagen..."
    wzOpenTitle = "Abrir"
    wzFile = String(255, Chr(0))
    wzInitialDir = CurrentProject.path 'Path de la BD
    wzFilter = "Imagen - JPEG " _
    & "(*.jpg)"
    wzFilterIndex = 1
    wzView = 1
    wzflags = 64
    wzfOpen = True

    ret = WizHook.GetFileName(wzhwndOwner, _
        wzAppName, wzDlgTitle, wzOpenTitle, wzFile, _
        wzInitialDir, wzFilter, wzFilterIndex, _
        wzView, wzflags, wzfOpen)
        
    ' Si no se ha pulsado el botón Cancelar (-302)
    If ret <> -302 Then
        EscogeImagen = wzFile
    Else
        EscogeImagen = ""
    End If
End Function

Public Function CargaImagenUsuario(Imagen As Image, idUsuario As Integer) As Boolean
    Dim Archivoimg As String
    Dim sexo As Integer
    
    sexo = DLookup("[fksexo]", "t_persona", "[id]=" & idUsuario)
    
    Archivoimg = rutaImagen & "foto_id" & Format(idUsuario, "000000") & ".jpg"
    
    If Dir(Archivoimg) = "" Then
        If (sexo = 1) Then
            Imagen.Picture = rutaImagen & "foto_man.jpg"
        Else
            Imagen.Picture = rutaImagen & "foto_woman.jpg"
        End If
        CargaImagenUsuario = False
    Else
        Imagen.Picture = Archivoimg
        CargaImagenUsuario = True
    End If
End Function

Public Function CopiarImagenUsuario(Imagen As Image, Ruta As String, idUsuario As Integer, Ancho As Integer, Alto As Integer) As Boolean
    Dim imgActual, Archivoimg As String
    Dim respuesta
  
    Archivoimg = rutaImagen & nombreImagenUsuario(idUsuario)
    imgActual = Imagen.Picture
    
    If (Ruta <> "") Then
        Imagen.Picture = Ruta
  
        If Imagen.ImageHeight > Alto Or Imagen.ImageWidth > Ancho Then
            Imagen.Picture = imgActual
            MsgBox "La imagén debe tener una tamaño máximo de 85x110 píxels."
            CopiarImagenUsuario = False
        Else
            If Dir(Archivoimg) <> "" Then
                respuesta = MsgBox("Va a sobreescribir la imagen existente por una nueva." & vbNewLine & "¿Desea continuar?", _
                           vbYesNo, "Alert: Imagen Usuario")
                If (respuesta = vbYes) Then
                    FileCopy Imagen.Picture, Archivoimg
                    Imagen.Picture = Archivoimg
                    CopiarImagenUsuario = True
                Else
                    Imagen.Picture = imgActual
                    CopiarImagenUsuario = False
                End If
            Else
                FileCopy Imagen.Picture, Archivoimg
                Imagen.Picture = Archivoimg
                CopiarImagenUsuario = True
            End If
        End If
    Else
        CopiarImagenUsuario = False
    End If
End Function

Public Function EliminarImagenUsuario(Ruta As String, idUsuario As Integer)
    Dim respuesta
    
    If (Ruta <> "" And InStr(Ruta, nombreImagenUsuario(idUsuario)) <> 0) Then
                respuesta = MsgBox("Va a eliminar la imagen existente sin posibilidad de recuperación." & vbNewLine & "¿Estas seguro?", _
                           vbYesNo, "Alert: Imagen Usuario")
                If (respuesta = vbYes) Then
                    Kill Ruta
                End If
    End If
End Function

Private Function nombreImagenUsuario(idUsuario As Integer) As String
    Dim nombreIMG As String

    nombreIMG = "foto_id" & Format(idUsuario, "000000") & ".jpg"
    nombreImagenUsuario = nombreIMG
End Function

