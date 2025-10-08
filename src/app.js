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
const formARoutes = require('./routes/formA');
const formTPARoutes = require('./routes/formTPA');
const formBRoutes = require('./routes/formB');
const sampleRoutes = require('./routes/samples');
const exportRoutes = require('./routes/export');

app.use('/', indexRoutes);
app.use('/form-a', formARoutes);
const formExtraRoutes = require('./routes/formExtras');
app.use('/form-tpa', formTPARoutes);
app.use('/form-b', formBRoutes);
app.use('/samples', sampleRoutes);
app.use('/export', exportRoutes);
app.use('/', formExtraRoutes);

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
