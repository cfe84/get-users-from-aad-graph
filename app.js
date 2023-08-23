const express = require('express');

function app(port) {
  const app = express();
  app.use(express.json());
  app.use(express.static('public'));

  app.get('/token', (req, res) => {
    if (req.headers["x-ms-token-aad-access-token"]) { 
      res.status(200).send(req.headers["x-ms-token-aad-access-token"]);
    } else {
      res.status(404);
    }
    res.end();
   });

  app.listen(port, () => {
    console.log(`Listening at http://localhost:${port}`);
  });
}

module.exports = app;