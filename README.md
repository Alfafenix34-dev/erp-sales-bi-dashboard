# ERP Sales & Coverage BI Dashboard

Sistema integral de automatizacion RPA e inteligencia comercial desarrollado en Python para extraer reportes diarios de ventas y devoluciones desde un ERP web, cruzarlos contra el rutero maestro de clientes y visualizar KPIs en tiempo real con control de acceso por roles.

## Stack Tecnologico
- **RPA y Scraping:** Python, Playwright (autenticacion automatizada, navegacion en menus dinamicos y manejo de calendario laboral/festivos en Colombia).
- **ETL y Analitica:** Pandas (cruce por llave compuesta Asesor + Cliente, separacion de canal organizacional vs. marca, calculo de cobertura, ticket promedio y drop size).
- **Backend y Seguridad:** Flask, Waitress, SQLite, hashes de contrasenas con Werkzeug, proteccion contra fuerza bruta y roles (Admin / Supervisor / Asesor).
- **Frontend y Despliegue:** HTML5, CSS3, JavaScript (Chart.js, filtros dinamicos, exportacion a Excel) y exposicion segura con Cloudflare Tunnel.

## Instalacion Rapida
```bash
pip install -r requirements.txt
playwright install chromium
copy .env.example .env
python dashboard_hb.py
```
