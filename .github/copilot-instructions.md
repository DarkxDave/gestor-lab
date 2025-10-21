# GitHub Copilot Instructions

## Project Context
- Stack: Node.js + Express + EJS
- DB: MySQL (via XAMPP/phpMyAdmin)
- Features: Multiple forms (textbox/checkbox), save/load sample IDs, export to Excel

## What to Generate
- Routes, controllers, models with async/await and prepared statements
- EJS views using Bootstrap
- Excel export logic
- Minimal, focused snippets aligned to the active file

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
- Express routers per feature
- Models: parameterized queries with mysql2
- Controllers: validate input, handle errors, return clear messages
- EJS: layouts/partials, Bootstrap classes
- Excel export via a maintained library (e.g., exceljs)
- Environment via .env (document required keys)

## Progress Tracking
- [x] copilot-instructions.md exists
- [x] Requirements clarified
- [x] Project scaffolded
- [x] Customized (forms/routes/models/export)
- [x] Extensions: none required
- [x] Dependencies installed
- [x] VS Code task: Run Dev Server (nodemon)
- [x] Launch at http://localhost:3000 with MySQL running
- [x] README updated (Windows/XAMPP)

## Task Completion Definition
- Project builds without errors
- This file and README are present and current
- Clear instructions to debug/launch are provided
