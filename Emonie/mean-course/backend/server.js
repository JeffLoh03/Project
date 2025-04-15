// Node.js server startup script//
const app = require('./app');
const debug = require('debug')('node-angular');
const http = require('http');

// The port value is normalized to ensure it matches the HTTP server's expectations
const normalizePort = val => {
  // Converting a string to a number, if cannot convert it will show 'NaN'
  var port = parseInt(val, 10);
  if (isNaN(port)) {
    return val;
  }
  if (port >= 0) {
    return port;
  }
  return false;
};

// Handling HTTP server startup errors
const onError = error => {
  if (error.syscall !== 'listen') {
    throw error;
  }
  // If the value is a string, it is displayed as "pipe", otherwise it is "port".
  const bind = typeof port === 'string' ? 'pipe ' + port : 'port ' + port;
  switch (error.code) {
    // If the port requires higher permissions, the user will be required to run the program with administrator permissions.
    case 'EACCES':
      console.error(bind + ' requires elevated privileges');
      process.exit(1);
      break;
    // If the port is already occupied, the computer will show that the port is already in use.
    case 'EADDRINUSE':
      console.error(bind + ' is already in use');
      process.exit(1);
      break;
    default:
      throw error;
  }
};

// It will be called after the server starts successfully.
const onListening = () => {
  // Get the address information of the current server.
  const addr = server.address();
  const bind = typeof port === 'string' ? 'pipe ' + port : 'port ' + port;
  debug('Listening on ' + bind);
  // It indicates that the server has been successfully running and is listening on the specified port.
  console.log('Server is running on port', port);
};

// If the environment variable is not set, the default value '3000' is used.
const port = normalizePort(process.env.PORT || '3000');
app.set('port', port);

// create a HTTP server
const server = http.createServer(app);
// When the server happen error, extract the function
server.on('error', onError);
server.on('listening', onListening);
// Start the server and start monitoring on the specified port
server.listen(port); 