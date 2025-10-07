# Gestor Lab

Aplicación web (Node.js + Express + EJS) para gestionar IDs de muestras con múltiples formularios, MySQL (phpMyAdmin) y exportación a Excel.

## Requisitos
- Windows con XAMPP (MySQL en `localhost:3306`)
- Node.js 18+

## Instalación
1. Clonar o abrir este proyecto en `c:\xampp\htdocs\gestor-lab`.
2. Copiar `.env.example` a `.env` y ajustar variables si aplica.
3. Instalar dependencias.

```powershell
# En PowerShell
cd c:\xampp\htdocs\gestor-lab
npm install
```

## Base de datos (phpMyAdmin)
1. Abrir phpMyAdmin (http://localhost/phpmyadmin).
2. Importar `scripts/init_db.sql` para crear la base `gestor_lab` y las tablas.

Alternativa por CLI (si tiene `mysql` en PATH):
```powershell
mysql -h localhost -P 3306 -u root -p < .\scripts\init_db.sql
```

## Ejecutar en desarrollo
```powershell
npm run dev
```

Abrir: http://localhost:3000

## Uso rápido
- Inicio: navegación a Formularios y Muestras.
- Formulario A y B: Ingrese `sample_id`, complete campos (textos y checkboxes) y "Guardar".
- Muestras: lista/búsqueda de `sample_id`; links para abrir en A/B.
- Exportar: `Exportar Excel` genera `muestras.xlsx` con datos A+B.

## Estructura
- `src/app.js`: servidor Express
- `src/db.js`: conexión MySQL (mysql2/promise)
- `src/routes/*`: rutas
- `src/controllers/*`: controladores
- `src/models/*`: consultas SQL
- `views/*`: EJS con Bootstrap
- `public/*`: estáticos
- `scripts/init_db.sql`: esquema MySQL

## Notas
- Usuario MySQL por defecto: `root` sin contraseña (ajuste `.env` si difiere).
- Asegúrese de que XAMPP MySQL esté iniciado.
- Para producción, configure variables de entorno seguras y un usuario MySQL propio.
