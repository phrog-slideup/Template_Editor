// Load environment variables from .env file
require('dotenv').config();

// Helper to get environment variable with default fallback
const getEnv = (key, defaultValue = '') => process.env[key] || defaultValue;

// Determine if we're in production mode
const isProduction = getEnv('NODE_ENV') === 'production';

// Get the appropriate base URL based on environment
const getBaseUrl = () => isProduction 
  ? getEnv('PRODUCTION_URL') 
  : getEnv('BASE_URL');

// Export configuration
const config = {
  // Environment
  nodeEnv: getEnv('NODE_ENV', 'development'),
  port: parseInt(getEnv('PORT', '3006'), 10),
  
  // URLs
  baseUrl: getBaseUrl(),
  
  // API
  apiPrefix: getEnv('API_PREFIX', '/api/slides'),
  
  // Upload paths
  uploadPath: getEnv('UPLOAD_PATH', '/node-editor/uploads'),
  uploadDir: getEnv('UPLOAD_DIR', 'app/uploads'),
  
  // File size limits
  maxFileSize: getEnv('MAX_FILE_SIZE', '100mb'),
  
  // Utility functions
  isProduction: isProduction,
  isDevelopment: !isProduction,
  
  // Generate full URLs for resources
  getResourceUrl: (resourcePath) => {
    // Remove any leading slash from resourcePath
    const cleanPath = resourcePath.startsWith('/') 
      ? resourcePath.substring(1) 
      : resourcePath;
    
    // For the upload URL, we need to handle the specific case
    if (cleanPath.startsWith('uploads/')) {
      // For uploads, use the site root, not the API path
      const baseWithoutApi = isProduction 
        ? 'https://kamaiipro.in/node-editor/'
        : 'http://localhost:3006/';
      
      return `${baseWithoutApi}${cleanPath}`;
    }
    
    // For other resources
    return `${getBaseUrl()}${cleanPath}`;
  }
};

module.exports = config;