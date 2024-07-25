function getUserInfoAsync(token, userId) {
  const options = {
    hostname: 'graph.microsoft.com',
    port: 443,
    path: `/v1.0/users/${userId}`,
    method: 'GET',
    headers: {
      Authorization: 'Bearer ' + token
    }
  }

  return new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      let data = "";
      if (res.statusCode != 200) {
        console.warn(userId, ' - statusCode:', res.statusCode);
      }
      res.on('data', (d) => {
        data += d;
      });
      res.on('end', () => {
        resolve(JSON.parse(data));
      });
    });
    req.on('error', (e) => {
      console.error(e);
      reject(e);
    });
    req.end();
  });
}

module.exports = getUserInfoAsync;