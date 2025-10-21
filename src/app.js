const express = require('express');
const path = require('path');
const morgan = require('morgan');
const dotenv = require('dotenv');
const ejsMate = require('ejs-mate');
dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

// Middlewares
app.use(morgan('dev'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, '..', 'public')));

// View engine
app.engine('ejs', ejsMate);
app.set('views', path.join(__dirname, '..', 'views'));
app.set('view engine', 'ejs');

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

app.use('/', indexRoutes);
app.use('/form-tpa', formTPARoutes);
app.use('/form-ram', formRAMRoutes);
app.use('/form-rmyl', formRMyLRoutes);
app.use('/form-ctcfe', formCTCFERoutes);
app.use('/form-sal', formSalRoutes);
app.use('/form-entero', formEnteroRoutes);
app.use('/form-saureus', formSaureusRoutes);
app.use('/samples', sampleRoutes);
app.use('/export', exportRoutes);

// 404
app.use((req, res) => {
  res.status(404).render('404', { title: 'No encontrado' });
});

// Error handler
app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).render('error', { title: 'Error', error: err });
});

app.listen(PORT, () => {
  console.log(`Servidor iniciado en http://localhost:${PORT}`);
});
