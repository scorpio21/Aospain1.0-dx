# Implementación de Cuentas - Servidor (Detallado)

El servidor debe gestionar la persistencia en archivos `.acc` y vincular la conexión con la cuenta activa.

## 1. Estructura del Archivo de Cuenta (`/Cuentas/USUARIO.acc`)
Crea una carpeta llamada `Cuentas` en el directorio raíz del servidor.
```ini
[INIT]
Password= (Hash MD5 de la contraseña)
Email=usuario@mail.com
PJ1=NombrePersonaje1
PJ2=NombrePersonaje2
PJ3=
... PJ10=
```

## 2. Lógica de Procesamiento (`HandleData` en `TCP.bas`)
Agrega este bloque antes de que el servidor verifique si el usuario está logueado:

```vb
' --- LOGIN DE CUENTA ---
If UCase$(Left$(rdata, 8)) = "LOGINACC" Then
    rdata = Right$(rdata, Len(rdata) - 8)
    Dim AccName As String, AccPass As String, AccPath As String
    AccName = UCase$(ReadField(1, rdata, 44))
    AccPass = ReadField(2, rdata, 44)
    AccPath = App.Path & "\Cuentas\" & AccName & ".acc"

    If Not FileExist(AccPath, vbNormal) Then
        Call SendData(ToIndex, UserIndex, 0, "ACCERRLa cuenta no existe.")
        Exit Sub
    End If

    If AccPass <> GetVar(AccPath, "INIT", "Password") Then
        Call SendData(ToIndex, UserIndex, 0, "ACCERRContraseña incorrecta.")
        Exit Sub
    End If

    ' Login exitoso: Enviamos la lista de PJs (Slots 1 al 10)
    Dim PjsList As String, j As Integer
    For j = 1 To 10
        PjsList = PjsList & GetVar(AccPath, "INIT", "PJ" & j) & ","
    Next j
    
    ' Guardamos el nombre de cuenta en la conexión del usuario temporalmente
    UserList(UserIndex).ConnID = AccName 
    Call SendData(ToIndex, UserIndex, 0, "ACCOK" & PjsList)
    Exit Sub
End If

' --- SELECCIÓN DE PERSONAJE ---
If UCase$(Left$(rdata, 7)) = "SELCHAR" Then
    Dim Slot As Integer, CharName As String
    Slot = val(Right$(rdata, Len(rdata) - 7))
    AccPath = App.Path & "\Cuentas\" & UserList(UserIndex).ConnID & ".acc"
    CharName = GetVar(AccPath, "INIT", "PJ" & Slot)

    If CharName <> "" Then
        ' Conectamos al personaje. 
        ' TIP: Puedes modificar ConnectUser para que no pida pass si viene de una cuenta.
        Call ConnectUser(UserIndex, CharName, "PASS_BYPASS") 
    Else
        ' Slot vacío: El cliente debe abrir el creador de personajes
        Call SendData(ToIndex, UserIndex, 0, "ABRIR_CREAR_PJ")
    End If
    Exit Sub
End If
```

## 3. Registro de Personaje Nuevo en Cuenta
Debes modificar la función `ConnectNewUser` o donde se crea el PJ para que, si el usuario tiene una cuenta activa (`UserList(UserIndex).ConnID`), se guarde el nombre del nuevo personaje en el primer slot libre del archivo `.acc`.
```vb
## 4. Sistema de Borrado con Confirmación
Agrega esta lógica a `HandleData` para gestionar la seguridad del borrado:

```vb
' --- SOLICITUD DE BORRADO ---
If UCase$(Left$(rdata, 10)) = "DELCHARREQ" Then
    Dim DelSlot As Integer
    DelSlot = val(Right$(rdata, Len(rdata) - 10))
    
    ' Generamos un código aleatorio de 6 dígitos
    UserList(UserIndex).Flags.DelCode = RandomNumber(100000, 999999)
    UserList(UserIndex).Flags.DelSlot = DelSlot
    
    ' Recuperamos el email de la cuenta
    Dim UserEmail As String
    UserEmail = GetVar(App.Path & "\Cuentas\" & UserList(UserIndex).ConnID & ".acc", "INIT", "Email")
    
    ' AQUÍ: Lógica para enviar el email (o simularlo en consola por ahora)
    Call SendData(ToIndex, UserIndex, 0, "||Se ha enviado un código a " & UserEmail & ". Úsalo para confirmar." & FONTTYPE_INFO)
    Debug.Print "CÓDIGO DE BORRADO PARA " & UserList(UserIndex).Name & ": " & UserList(UserIndex).Flags.DelCode
    
    Call SendData(ToIndex, UserIndex, 0, "ASKDELCODE")
    Exit Sub
End If

' --- CONFIRMACIÓN DE BORRADO ---
If UCase$(Left$(rdata, 11)) = "DELCHARCONF" Then
    Dim ConfCode As String
    ConfCode = Right$(rdata, Len(rdata) - 11)
    
    If ConfCode = CStr(UserList(UserIndex).Flags.DelCode) Then
        ' El código coincide, procedemos al borrado
        AccPath = App.Path & "\Cuentas\" & UserList(UserIndex).ConnID & ".acc"
        CharName = GetVar(AccPath, "INIT", "PJ" & UserList(UserIndex).Flags.DelSlot)
        
        ' 1. Borramos el nombre del slot en la cuenta
        Call WriteVar(AccPath, "INIT", "PJ" & UserList(UserIndex).Flags.DelSlot, "")
        
        ' 2. Borramos el archivo del personaje (.chr)
        If CharName <> "" Then
            If FileExist(App.Path & "\charfile\" & CharName & ".chr", vbNormal) Then
                Kill App.Path & "\charfile\" & CharName & ".chr"
            End If
        End If
        
        Call SendData(ToIndex, UserIndex, 0, "DELOK")
        ' Limpiamos los flags de seguridad
        UserList(UserIndex).Flags.DelCode = 0
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Código de confirmación incorrecto." & FONTTYPE_INFO)
    End If
    Exit Sub
End If
```
