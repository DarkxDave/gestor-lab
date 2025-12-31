const express = require('express');
const path = require('path');
const morgan = require('morgan');
const dotenv = require('dotenv');
// const ejsMate = require('ejs-mate');
dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

// Middlewares
app.use(morgan('dev'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
// Static assets are served by the Angular app during development
// app.use(express.static(path.join(__dirname, '..', 'public')));

// View engine
// Disable EJS view engine; this server is a pure JSON API now
// app.engine('ejs', ejsMate);
// app.set('views', path.join(__dirname, '..', 'views'));
// app.set('view engine', 'ejs');

// Routes
const indexRoutes = require('./routes/index');
const formTPARoutes = require('./routes/formTPA');
const formRAMRoutes = require('./routes/formRAM');
const formRMyLRoutes = require('./routes/formRMyL');
const formCTCFERoutes = require('./routes/formCTCFE');
const formSalRoutes = require('./routes/formSal');
const formEnteroRoutes = require('./routes/formEntero');
const formSaureusRoutes = require('./routes/formSaureus');
const sampleRoutes = require('./routes/samples');
const exportRoutes = require('./routes/export');

// Prefix all routes with /api for frontend proxying
app.use('/api', indexRoutes);
app.use('/api/form-tpa', formTPARoutes);
app.use('/api/form-ram', formRAMRoutes);
app.use('/api/form-rmyl', formRMyLRoutes);
app.use('/api/form-ctcfe', formCTCFERoutes);
app.use('/api/form-sal', formSalRoutes);
app.use('/api/form-entero', formEnteroRoutes);
app.use('/api/form-saureus', formSaureusRoutes);
app.use('/api/samples', sampleRoutes);
app.use('/api/export', exportRoutes);

// 404
app.use((req, res) => {
  res.status(404).json({ message: 'No encontrado' });
});

// Error handler
app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).json({ message: 'Error interno', error: err.message });
});

app.listen(PORT, () => {
  console.log(`Servidor iniciado en http://localhost:${PORT}`);
});
