Option Explicit

Dim filePath
filePath = "C:\Users\User\Desktop\prueba.txt"

' Función para mostrar una alerta y obtener la respuesta del usuario
Function ConfirmAction(message)
    Dim objShell, response
    Set objShell = CreateObject("WScript.Shell")
    response = objShell.Popup(message, 10, "Confirmar acción", 4 + 32)
    Set objShell = Nothing
    ConfirmAction = (response = 6) ' 6 = Botón "Sí" en el mensaje
End Function

' Verificar si el archivo existe antes de proceder
Dim FileSystemObject
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
If FileSystemObject.FileExists(filePath) Then
    ' Mostrar el mensaje de confirmación y obtener la respuesta del usuario
    If ConfirmAction("¿Estás seguro que deseas borrar el archivo " & filePath & "?") Then
        ' Si el usuario confirmó, proceder a eliminar el archivo
        FileSystemObject.DeleteFile(filePath)
        MsgBox "El archivo se ha eliminado correctamente.", vbInformation, "Eliminación exitosa"
    Else
        ' Si el usuario no confirmó, mostrar un mensaje de cancelación
        MsgBox "El archivo no ha sido eliminado.", vbExclamation, "Operación cancelada"
    End If
Else
    ' Si el archivo no existe, mostrar un mensaje de error
    MsgBox "El archivo " & filePath & " no existe.", vbCritical, "Error"
End If

Set FileSystemObject = Nothing
