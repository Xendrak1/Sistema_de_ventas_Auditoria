# Vulnerabilidades Intencionales

Este archivo documenta las fallas incluidas de forma intencional para fines academicos. La finalidad es permitir un analisis de seguridad, generar evidencias y proponer controles.

## 1. Credenciales debiles

Usuarios de prueba:

- `admin / admin123`
- `ventas / ventas123`

Riesgo: acceso no autorizado por fuerza bruta o adivinacion de contrasenas.

Control recomendado: exigir contrasenas robustas, bloqueo por intentos fallidos y cambio periodico.

## 2. Contrasenas en texto plano

La tabla `usuarios` almacena las contrasenas sin hash.

Riesgo: si la base de datos es expuesta, las credenciales quedan comprometidas inmediatamente.

Control recomendado: almacenar contrasenas con `bcrypt`, `argon2` o un mecanismo equivalente.

## 3. SQL injection en login

Ruta afectada:

```text
/login
```

Prueba manual:

```text
Usuario: ' OR '1'='1' --
Contrasena: cualquier valor
```

Riesgo: ingreso sin credenciales validas.

Control recomendado: usar consultas parametrizadas en todas las operaciones SQL.

## 4. SQL injection en busqueda

Ruta afectada:

```text
/buscar?q=
```

Riesgo: manipulacion de consultas y exposicion no autorizada de datos.

Control recomendado: usar parametros SQL y validar la entrada del usuario.

## 5. XSS reflejado

Ruta afectada:

```text
/buscar?q=<script>alert(1)</script>
```

Riesgo: ejecucion de codigo JavaScript en el navegador de un usuario autenticado.

Control recomendado: escapar salidas HTML y evitar `render_template_string` con datos concatenados.

## 6. XSS almacenado

La descripcion de productos se muestra usando `safe`.

Prueba:

```html
<script>alert("producto")</script>
```

Riesgo: ejecucion persistente de JavaScript al listar productos.

Control recomendado: sanitizar entradas y no marcar contenido de usuario como seguro.

## 7. Falta de CSRF

Los formularios no incluyen tokens CSRF.

Riesgo: un atacante podria inducir acciones no autorizadas desde el navegador de un usuario autenticado.

Control recomendado: implementar proteccion CSRF con Flask-WTF o mecanismo similar.

## 8. Clave de sesion debil

La aplicacion usa una `SECRET_KEY` fija y predecible.

Riesgo: manipulacion de cookies de sesion si la clave es descubierta.

Control recomendado: usar una clave fuerte desde variables de entorno.

## 9. Modo debug activo

La aplicacion se ejecuta con `debug=True`.

Riesgo: exposicion de informacion tecnica y consola interactiva en configuraciones inseguras.

Control recomendado: desactivar debug en entornos desplegados.

## 10. Cabeceras de seguridad ausentes

La aplicacion no configura cabeceras como `Content-Security-Policy`, `X-Frame-Options` o `Strict-Transport-Security`.

Riesgo: mayor exposicion a XSS, clickjacking y ataques relacionados.

Control recomendado: agregar cabeceras de seguridad con middleware o configuracion del servidor web.
