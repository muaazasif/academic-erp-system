// server.js - Proxy server to run Python Flask app on Glitch
const express = require('express');
const { spawn } = require('child_process');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Serve static files
app.use('/static', express.static('static'));

// Proxy requests to the Python Flask app
let pythonProcess;

app.get('*', (req, res) => {
  // For Glitch, we'll serve a simple message indicating the Python app should be running
  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>ERP System Loading...</title>
      <style>
        body { font-family: Arial, sans-serif; text-align: center; padding: 50px; background-color: #f0f0f0; }
        .loader { border: 16px solid #f3f3f3; border-radius: 50%; border-top: 16px solid #3498db; width: 120px; height: 120px; animation: spin 2s linear infinite; margin: 0 auto; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
      </style>
    </head>
    <body>
      <div class="loader"></div>
      <h2>ERP System is starting up...</h2>
      <p>If you're seeing this page, the Python Flask application is starting. This may take a moment.</p>
      <p>Once started, you can access the application at the main URL.</p>
    </body>
    </html>
  `);
});

// Start the Python Flask app as a child process
function startPythonApp() {
  console.log('Starting Python Flask application...');
  
  pythonProcess = spawn('python3', ['app.py'], {
    cwd: __dirname,
    env: {
      ...process.env,
      PORT: PORT,
      FLASK_ENV: 'production'
    }
  });

  pythonProcess.stdout.on('data', (data) => {
    console.log(`Python app stdout: ${data}`);
  });

  pythonProcess.stderr.on('data', (data) => {
    console.error(`Python app stderr: ${data}`);
  });

  pythonProcess.on('close', (code) => {
    console.log(`Python app exited with code ${code}`);
    // Restart the Python app after a delay
    setTimeout(startPythonApp, 5000);
  });
}

// Start the Express server
app.listen(PORT, () => {
  console.log(`Proxy server listening on port ${PORT}`);
  // Start the Python Flask app
  startPythonApp();
});