const express = require('express');
const cors = require('cors');
const path = require('path'); // 🔥 ADDED THIS
require('dotenv').config();

const dataRoutes = require('./routes/dataRoutes');
const cronRoutes = require('./routes/cronRoutes');
const { initCron } = require('./controllers/cronController');

const app = express();

app.use((req, res, next) => {
  console.log('HIT:', req.method, req.url, new Date().toISOString());
  next();
});

app.use(cors({
    origin: true,
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'Accept']
}));

app.use(express.json());

// 1. API ROUTES MUST GO FIRST
app.use('/api/data', dataRoutes);
app.use('/api/cron', cronRoutes);

// 2. SERVE THE REACT PRODUCTION BUILD
// This points to the build folder you just created in the client directory
const buildPath = path.join(__dirname, '../client/build');
app.use(express.static(buildPath));

// 3. CATCH-ALL ROUTE FOR REACT ROUTER
// If someone visits /login or /dashboard directly, send them the React app
app.get(/.*/, (req, res) => {
    res.sendFile(path.join(buildPath, 'index.html'));
});

const PORT = process.env.PORT || 5000;

app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on port ${PORT}`);
});
app.listen(PORT, async () => {
    console.log(`Server running on port ${PORT}`);
    await initCron(); // Initialize background cron!
});