# Guia de Despliegue y Auditoria

Proyecto: Auditoria de Seguridad y Politicas de Contrasenas para un Sistema de Ventas en Fedora 39.

## 1. Objetivo del despliegue

Implementar un sistema web de ventas en un servidor Fedora 39 dentro de VirtualBox para realizar una auditoria de seguridad con herramientas como Nmap y Nessus.

## 2. Herramientas utilizadas

| Herramienta | Uso dentro del proyecto |
| --- | --- |
| VirtualBox | Creacion del entorno de laboratorio |
| Fedora 39 | Sistema operativo del servidor |
| Python 3 | Ejecucion del sistema web |
| Flask | Framework web del sistema de ventas |
| SQLite | Base de datos local de prueba |
| Git y GitHub | Control de versiones y entrega del codigo |
| Nmap | Escaneo de puertos y servicios |
| Nessus Essentials | Analisis de vulnerabilidades |
| Word | Elaboracion del informe con normas APA |

## 3. Pasos de despliegue en Fedora

```bash
sudo dnf update -y
sudo dnf install -y git python3 python3-pip
git clone https://github.com/Xendrak1/Sistema_de_ventas_Auditoria.git
cd Sistema_de_ventas_Auditoria
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

El servicio web quedara disponible en:

```text
http://IP_DEL_SERVIDOR:5000
```

## 4. Configuracion de red sugerida

En VirtualBox se recomienda usar una de estas opciones:

| Modo de red | Uso recomendado |
| --- | --- |
| Adaptador puente | Permite escanear la VM desde otra maquina de la red |
| Red interna | Permite laboratorio aislado entre VMs |
| NAT con reenvio de puerto | Permite pruebas basicas desde el equipo anfitrion |

Para una auditoria con Nmap y Nessus, el modo mas simple suele ser Adaptador puente.

## 5. Comandos de verificacion en Fedora

```bash
ip addr
ss -tulnp
sudo firewall-cmd --list-all
```

Si el puerto 5000 no es accesible:

```bash
sudo firewall-cmd --add-port=5000/tcp --permanent
sudo firewall-cmd --reload
```

## 6. Analisis con Nmap

Reemplazar `IP_DEL_SERVIDOR` por la IP real de Fedora:

```bash
nmap -sV -p 5000 IP_DEL_SERVIDOR
nmap -A IP_DEL_SERVIDOR
nmap -p- IP_DEL_SERVIDOR
```

Evidencias recomendadas:

- Captura del comando ejecutado.
- Captura del puerto 5000 abierto.
- Version del servicio detectado.
- Puertos adicionales encontrados en Fedora.

## 7. Analisis con Nessus

Pasos:

1. Crear un escaneo de tipo `Basic Network Scan`.
2. Colocar como objetivo la IP de Fedora.
3. Ejecutar el escaneo.
4. Exportar el reporte en PDF o HTML.
5. Tomar capturas de vulnerabilidades, severidad y recomendaciones.

Evidencias recomendadas:

- Pantalla de configuracion del scan.
- Pantalla del scan finalizado.
- Resumen de vulnerabilidades.
- Detalle de vulnerabilidades relevantes.

## 8. Vulnerabilidades esperadas del sistema web

| Vulnerabilidad | Ubicacion | Riesgo |
| --- | --- | --- |
| Credenciales debiles | Login | Acceso no autorizado |
| Contrasenas en texto plano | Base SQLite | Exposicion de credenciales |
| SQL injection | `/login` y `/buscar` | Manipulacion de consultas |
| XSS reflejado | `/buscar?q=` | Ejecucion de JavaScript |
| XSS almacenado | Descripcion de productos | Ejecucion persistente de JavaScript |
| Falta de CSRF | Formularios | Acciones no autorizadas |
| Debug activo | Flask | Exposicion de informacion tecnica |
| Cabeceras ausentes | Respuestas HTTP | Mayor superficie de ataque |

## 9. Tabla base de analisis de riesgos

| Activo | Amenaza | Vulnerabilidad | Impacto | Probabilidad | Nivel de riesgo | Control recomendado |
| --- | --- | --- | --- | --- | --- | --- |
| Sistema web | Acceso no autorizado | Credenciales debiles | Alto | Alta | Alto | Politica de contrasenas y bloqueo por intentos |
| Sistema web | Manipulacion de consultas | SQL injection | Alto | Media | Alto | Consultas parametrizadas |
| Sistema web | Robo de sesion o ejecucion de scripts | XSS | Medio | Media | Medio | Escapar salidas y sanitizar entradas |
| Base de datos | Exposicion de credenciales | Contrasenas en texto plano | Alto | Media | Alto | Hash seguro con bcrypt o argon2 |
| Servidor Fedora | Enumeracion de servicios | Puertos abiertos | Medio | Alta | Alto | Firewall y cierre de servicios innecesarios |
| Aplicacion Flask | Exposicion tecnica | Debug activo | Alto | Media | Alto | Desactivar debug en despliegue |

## 10. Capturas que deben incluirse en el informe

- VirtualBox con Fedora 39 funcionando.
- IP del servidor obtenida con `ip addr`.
- Sistema web abierto en navegador.
- Login del sistema.
- Panel principal del sistema.
- Resultado de Nmap.
- Resultado de Nessus.
- Tabla de riesgos.
- Controles propuestos.
