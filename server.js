const express = require('express');
const cors = require('cors');
const path = require('path');
const os = require('os');
const { spawn } = require('child_process');

const app = express();
const PORT = 8080;
const FLASK_PORT = 5000;

// Configurar CORS para permitir acceso desde cualquier origen
app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'PUT', 'DELETE'],
    allowedHeaders: ['Content-Type', 'Authorization']
}));

// Servir archivos estáticos desde la carpeta static
app.use('/static', express.static(path.join(__dirname, 'static')));
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));
app.use(express.json());

// Función para obtener la IP local
function getLocalIP() {
    const interfaces = os.networkInterfaces();
    
    // Priorizar adaptadores WiFi
    const wifiInterfaces = ['Wi-Fi', 'Wireless', 'wlan', 'wifi'];
    
    for (const name of Object.keys(interfaces)) {
        // Buscar primero adaptadores WiFi
        if (wifiInterfaces.some(wifi => name.toLowerCase().includes(wifi.toLowerCase()))) {
            for (const interface of interfaces[name]) {
                if (interface.family === 'IPv4' && !interface.internal) {
                    return interface.address;
                }
            }
        }
    }
    
    // Si no encuentra WiFi, buscar cualquier adaptador que no sea virtual
    for (const name of Object.keys(interfaces)) {
        // Evitar adaptadores virtuales
        if (!name.toLowerCase().includes('virtual') && 
            !name.toLowerCase().includes('vmware') && 
            !name.toLowerCase().includes('vbox') &&
            !name.toLowerCase().includes('docker') &&
            !name.toLowerCase().includes('wsl')) {
            
            for (const interface of interfaces[name]) {
                if (interface.family === 'IPv4' && !interface.internal) {
                    return interface.address;
                }
            }
        }
    }
    
    return 'localhost';
}

// Función para iniciar el servidor Flask
function iniciarFlask() {
    console.log('🚀 Iniciando servidor Flask...');
    
    const flaskProcess = spawn('python', ['app.py'], {
        cwd: __dirname,
        stdio: ['pipe', 'pipe', 'pipe']
    });

    flaskProcess.stdout.on('data', (data) => {
        console.log(`Flask: ${data.toString()}`);
    });

    flaskProcess.stderr.on('data', (data) => {
        console.error(`Flask Error: ${data.toString()}`);
    });

    flaskProcess.on('close', (code) => {
        console.log(`Flask se cerró con código: ${code}`);
    });

    return flaskProcess;
}

// Proxy para redirigir todas las peticiones a Flask
app.use('*', (req, res) => {
    const targetUrl = `http://localhost:${FLASK_PORT}${req.originalUrl}`;
    
    // Redirigir a Flask
    res.redirect(targetUrl);
});

// API para obtener información del servidor
app.get('/api/server-info', (req, res) => {
    const localIP = getLocalIP();
    res.json({
        nombre: 'Servidor Llenado Automático CDMX-GDL',
        version: '1.0.0',
        puerto: PORT,
        flaskPuerto: FLASK_PORT,
        ip: localIP,
        urls: {
            local: `http://localhost:${PORT}`,
            redLocal: `http://${localIP}:${PORT}`,
            flask: `http://localhost:${FLASK_PORT}`,
            internet: 'https://[URL-DE-LOCALTUNNEL]'
        },
        estado: 'Activo',
        timestamp: new Date().toISOString()
    });
});

// Iniciar servidor
app.listen(PORT, '0.0.0.0', () => {
    const localIP = getLocalIP();
    console.log('🚀 Servidor Llenado Automático CDMX-GDL iniciado!');
    console.log('📍 Información de acceso:');
    console.log(`   Local: http://localhost:${PORT}`);
    console.log(`   Red Local: http://${localIP}:${PORT}`);
    console.log(`   Puerto: ${PORT}`);
    console.log(`   Flask: http://localhost:${FLASK_PORT}`);
    console.log('📋 Para acceder desde otras computadoras en la red:');
    console.log(`   1. Asegúrate de que el firewall permita conexiones al puerto ${PORT}`);
    console.log(`   2. Usa la IP: http://${localIP}:${PORT}`);
    console.log('🔧 Para desarrollo, usa: npm run dev');
    console.log('⏹️  Para detener el servidor: Ctrl+C');
    
    // Iniciar Flask después de un breve delay
    setTimeout(() => {
        iniciarFlask();
    }, 1000);
});

// Manejar errores no capturados
process.on('uncaughtException', (error) => {
    console.error('Error no capturado:', error);
});

process.on('unhandledRejection', (reason, promise) => {
    console.error('Promesa rechazada no manejada:', reason);
}); 