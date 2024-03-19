const fs = require('fs');
const https = require('https');
const readline = require('node:readline');

const UNAUTHORIZED = "Unauthorized";
const TEST_USER_ID = "2a38ca76-123d-4f5d-b139-a70daed5c6c0";
const TOKEN_FILE = ".token";

Promise.parallel = function(degree, promises) {
  let index = 0;
  const results = [];
  const executors = [];
  for(let executorId = 0; executorId < degree; executorId++) {
    executors.push((async () => {
      let i;
      while ((i = index++) < promises.length) {
        results[i] = await promises[i]();
      }
    })());
  }
  return Promise.allSettled(executors)
    .then(() => results);
}

function readFromConsoleAsync(message) {
  const { stdin: input, stderr: output } = require('node:process');
  const rl = readline.createInterface({ input, output });

  return new Promise(resolve => {
    rl.question(message, (answer) => {
      rl.close();
      resolve(answer);
    });
  })
}

function getUsers(file) {
  const content = fs.readFileSync(file);
  const lines = content.toString().split('\n');
  const headers = lines[0].trim().split(',');
  const users = [];
  for(let i = 1; i < lines.length; i++) {
    const line = lines[i].replace("\r", "").split(',');
    const obj = {};
    for(let j = 0; j < headers.length; j++) {
      obj[headers[j]] = line[j];
    }
    users.push(obj);
  }
  return {users, headers};
}

function getUserInfoAsync(userId, token) {
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
      console.warn(userId, ' - statusCode:', res.statusCode);
      if (res.statusCode === 401) {
        reject(UNAUTHORIZED);
      }
      let data = "";
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

function getUserId(user) {
  return user.organizer_id || user.user_id;
}

async function refreshTokenAsync() {
  token = await readFromConsoleAsync(`Token is dead. Please refresh go to https://developer.microsoft.com/en-us/graph/graph-explorer, and provide it: `);
  saveToken(token);
  return token;
}

async function getTokenAsync() {
  let token = "";
  if (fs.existsSync(TOKEN_FILE)) {
    token = fs.readFileSync(TOKEN_FILE).toString();
    try {
      await getUserInfoAsync(TEST_USER_ID, token);
    } catch(e) {
      token = refreshTokenAsync();
    }
  } else {
    token = refreshTokenAsync();
  }
  return token;
}

function saveToken(token) {
  fs.writeFileSync(TOKEN_FILE, token);
}

async function enrichUser(user, token) {
  const data = await getUserInfoAsync(user.organizer_id, token);
  user["displayName"] = data["displayName"];
  user["givenName"] = data["givenName"];
  user["jobTitle"] = data["jobTitle"];
  user["mail"] = data["mail"];
}

function displayUsers(users) {
  const headers = users.headers.concat(["displayName", "givenName", "jobTitle", "mail"])
  console.log(headers.join(","));
  for(let user of users.users) {
    const line = headers.map(header => user[header] || "").join(",")
    console.log(line);
  }
}

async function runAsync(file) {
  let token = await getTokenAsync();
  const users = getUsers(file);
  await enrichUsers(users, token);
  displayUsers(users);
}

async function enrichUsers(users, token) {
  const promises = users.users.map(user => () => enrichUser(user, token));
  await Promise.parallel(20, promises);
}

const file = process.argv[2];

if (!file) {
  console.error("Please specify a file name");
  return;
}
runAsync(file).then();


