# Gestor Lab

Aplicación web (Node.js + Express + EJS) para gestionar IDs de muestras con múltiples formularios, MySQL (phpMyAdmin) y exportación a Excel.

## Tecnologías utilizadas

### Frontend
- Angular
- Ionic Framework
- TypeScript
- HTML5
- CSS

### Backend
- Node.js
- ExcelJS

## Requisitos
- Windows
- XAMPP (MySQL en `localhost:3306`)
- Node.js 18+
- Angular CLI
- Ionic CLI

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
2. Importar `scripts/init_db.sql`.

Alternativa por CLI (si tiene `mysql` en PATH):
```powershell
mysql -h localhost -P 3306 -u root -p < .\scripts\init_db.sql
```

### Sin migraciones incrementales
Este proyecto se instala limpio con un único script baseline (`init_db.sql`). 

## Ejecutar en desarrollo
```powershell
npm run dev
```

Abrir: http://localhost:3000

### Ejecutar el frontend (Angular + Ionic)
Ejecutar el frontend (Angular + Ionic)
cd frontend
ionic serve



## Uso rápido
- Inicio: navegación a Formularios y Muestras.
- Formulario A y B: Ingrese `sample_id`, complete campos (textos y checkboxes) y "Guardar".
- Muestras: lista/búsqueda de `sample_id`; links para abrir en A/B.
- Exportar: `Exportar Excel` genera `muestras.xlsx` con datos A+B.

## Estructura
Estructura del proyecto
### Backend

- src/app.js: servidor Express
- src/db.js: conexión a MySQL
- src/routes/: definición de rutas
- src/controllers/: lógica de negocio
- src/models/: consultas SQL
- scripts/init_db.sql: esquema de base de datos

### Frontend
- frontend/src/app/: componentes Angular
- frontend/src/services/: servicios de comunicación con el backend
- frontend/src/assets/: recursos estáticos
- frontend/src/theme/: estilos

## Notas
- Usuario MySQL por defecto: `root` sin contraseña (ajuste `.env` si difiere).
- Asegúrese de que XAMPP MySQL esté iniciado.
- Para producción, configure variables de entorno seguras y un usuario MySQL propio.
