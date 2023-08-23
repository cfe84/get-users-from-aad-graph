const express = require('express');

function app(port) {
  const app = express();
  app.use(express.json());
  app.use(express.static('public'));

  app.get('/token', (req, res) => {
    if (req.headers["X-MS-TOKEN-AAD-ACCESS-TOKEN"] || req.headers["X-MS-TOKEN-AAD-ID-TOKEN"]) { 
      res.status(200).send(req.headers["X-MS-TOKEN-AAD-ACCESS-TOKEN"]);
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