# GitHub Copilot Instructions

## Project Context
- **Frontend**: Ionic/Angular (Standalone Components, TypeScript, SCSS) running on port 8100
- **Backend**: Node.js + Express (JSON API only, no EJS views) running on port 3000
- **DB**: MySQL (via XAMPP/phpMyAdmin)
- **Architecture**: Separate frontend SPA and backend REST API; frontend proxies `/api/*` to backend
- **Features**: Multiple forms (textbox/checkbox), create/edit samples, load/save via API, export to Excel

## What to Generate
**Backend**: Routes, controllers, models with async/await and prepared statements; respond with JSON (no EJS rendering)
**Frontend**: Standalone Angular components (with imports: IonicModule, FormsModule, CommonModule, RouterModule, HttpClientModule), services for API calls, Bootstrap/Ionic styling
**Shared**: Excel export logic (via exceljs), HTTP client patterns for `/api/*` endpoints

## Constraints and Rules
- Working directory: "."
- No external links unless requested
- Use placeholders only when necessary; clearly mark them for replacement
- Do not install extensions unless explicitly specified
- Do not create new folders except .vscode for tasks.json
- Keep explanations concise; avoid verbose output
- If a feature is not confirmed, ask for clarification first
- For VS Code: assume integrated terminal, output pane, unit tests, and tasks
- Use Node + mysql2, dotenv for secrets, nodemon for dev

## Development Conventions
**Backend (Express API)**:
- All routes prefixed with `/api` (e.g., `/api/form-tpa`, `/api/samples`, `/api/export/all`)
- Controllers: validate input, handle errors, return JSON responses (e.g., `res.json({ data, message })` or `res.status(500).json({ error })`)
- Models: parameterized queries with mysql2; no EJS view rendering
- No static view engine; disabled `app.set('view engine', 'ejs')` and view directory

**Frontend (Ionic/Angular)**:
- Standalone components (each imports dependencies it uses)
- Inject `HttpClient`, `Router`, `ActivatedRoute` as needed
- Call `/api/*` endpoints (frontend dev server proxies via `proxy.conf.json` to `http://localhost:3000`)
- Routes: add new pages in `src/app/pages/*` and register in `app.routes.ts`
- Use `ActionSheetController`, `IonSegment`, etc. from `@ionic/angular`

**Shared**:
- Excel export: use exceljs, triggered via `/api/export/all?sample_id=...`
- Environment: `.env` in backend only (API credentials, DB config)

## Progress Tracking
- [x] Ionic/Angular frontend scaffolded (standalone components)
- [x] Backend converted to JSON API (no EJS)
- [x] All routes prefixed with `/api`
- [x] Frontend-backend communication via HTTP/proxy
- [x] Create/edit sample forms working
- [x] Excel export logic integrated
- [x] MySQL integration via mysql2
- [x] Development servers: backend on 3000, frontend on 8100
- [x] nodemon watching backend changes

## Running the Application
**Backend**: `cd ionic-app/server && npm run dev` (listens on `http://localhost:3000`)
**Frontend**: `cd ionic-app && ionic serve --host=localhost --port=8100 --no-open` (listens on `http://localhost:8100`)
**Database**: Ensure MySQL is running via XAMPP; tables created via `server/scripts/init_db.sql`

## Task Completion Definition
- Backend API starts without errors (`Servidor iniciado en http://localhost:3000`)
- Frontend compiles without errors (ng serve ready)
- Forms load, create/edit samples, save via API, export to Excel
- Router navigation stable; no unexpected redirects
- All controllers return JSON; no EJS view rendering
- This file and README updated and current
