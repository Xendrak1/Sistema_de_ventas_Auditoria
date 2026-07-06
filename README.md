# Sistema de Ventas para Auditoria

Aplicacion web sencilla desarrollada con Python y Flask para un proyecto academico de auditoria de seguridad. El sistema simula la gestion basica de productos, clientes y ventas en un entorno Linux Fedora 39.

> Importante: este proyecto contiene vulnerabilidades intencionales para realizar pruebas controladas con Nmap, Nessus y analisis manual. No debe usarse en produccion.

## Funcionalidades

- Inicio de sesion.
- Panel principal con metricas.
- Registro y listado de productos.
- Productos con SKU, categoria, stock e imagen local.
- Registro y listado de clientes.
- Registro y listado de ventas.
- Reportes basicos de ventas e inventario.
- Alertas visuales de bajo stock.
- Busqueda de productos.

## Credenciales de prueba

```text
Usuario: admin
Contrasena: admin123
```

```text
Usuario: ventas
Contrasena: ventas123
```

## Requisitos

- Fedora 39 o superior.
- Python 3.
- Git.
- Navegador web.

## Instalacion en Fedora

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

Abrir desde el navegador:

```text
http://IP_DEL_SERVIDOR:5000
```

Para obtener la IP del servidor:

```bash
ip addr
```

## Habilitar acceso desde otra maquina

Si el firewall de Fedora esta activo, abrir el puerto del laboratorio:

```bash
sudo firewall-cmd --add-port=5000/tcp --permanent
sudo firewall-cmd --reload
sudo firewall-cmd --list-all
```

## Pruebas sugeridas con Nmap

Desde otra maquina de la misma red:

```bash
nmap -sV -p 5000 IP_DEL_SERVIDOR
nmap -A IP_DEL_SERVIDOR
nmap -p- IP_DEL_SERVIDOR
```

## Pruebas sugeridas con Nessus

1. Crear un nuevo scan de tipo Basic Network Scan.
2. Colocar como target la IP del servidor Fedora.
3. Ejecutar el escaneo.
4. Exportar el reporte en PDF o HTML.
5. Tomar capturas del panel de resultados y vulnerabilidades encontradas.

## Estructura

```text
.
├── app.py
├── requirements.txt
├── static/
│   ├── img/
│   └── styles.css
├── templates/
└── VULNERABILIDADES.md
```

## Nota para el informe

El sistema fue creado para representar un activo auditable dentro del proyecto: "Auditoria de Seguridad y Politicas de Contrasenas para un Sistema de Ventas en un entorno Linux-Fedora 39".
