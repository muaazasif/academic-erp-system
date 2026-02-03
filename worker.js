// worker.js - Cloudflare Worker to serve Flask application
import app from './app.js';

export default {
  async fetch(request, env, ctx) {
    // Modify the request to work with the Flask app
    const url = new URL(request.url);
    
    // Handle static files
    if (url.pathname.startsWith('/static/')) {
      // Return static files directly
      const staticResponse = await fetch(`https://your-bucket-name.s3.amazonaws.com${url.pathname}`);
      return staticResponse;
    }
    
    // For all other requests, pass to Flask app
    return app.fetch(request, env, ctx);
  },
};