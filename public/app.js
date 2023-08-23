async function getUserInfoAsync(token, userId) {
  const options = {
    headers: {
      Authorization: 'Bearer ' + token
    }
  };

  const res = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}`, options);
  if (res.status != 200) {
    return null;
  }
  const data = await res.json();
  return data;
}

async function getTokenAsync() {
  const res = await fetch('/token');
  if (res.status != 200) {
    return null;
  }
  const data = await res.text();
  return data;
}

const fields = ["id", "displayName", "mail", "jobTitle", "officeLocation", "mobilePhone", "businessPhones"];

function userInfoToRow(fields, userInfo) {
  const columns = fields.map(field => `<td>${!!userInfo ? userInfo[field] : "Not found"}\t</td>`);
  return `<tr>${columns.join('')}</tr>`;
}

function titleRow(fields) {
  const columns = fields.map(field => `<th>${field}\t</th>`);
  return `<tr>${columns.join('')}</tr>`;
}

let cache = {};
let aadtoken = null;

async function processAsync() {
  const token = aadtoken || document.getElementById('input-token').value;
  const output = document.getElementById('output');
  const inputTextArea = document.getElementById('input');
  const input = inputTextArea.value;
  const inputs = input.split('\n');
  let res = "";
  let i = 1;
  for(let input of inputs) {
    output.innerHTML = `<tbody>${titleRow(fields)}${res}</tbody><tr><td colspan="6">Processing ${i++} of ${inputs.length}</td></tr>`;
    if (cache[input] === undefined) {
      console.log('cache miss');
      const userInfo = await getUserInfoAsync(token, input);
      cache[input] = userInfo;
    } else {
      console.log('cache hit');
    }
    res += `${userInfoToRow(fields, cache[input])}`;
  }
  res = `<tbody>${titleRow(fields)}${res}</tbody>`;
  output.innerHTML = res;
}

window.onload = async () => {
  const button = document.getElementById('btn-getinfo');
  button.onclick = () => processAsync().then();

  aadtoken = await getTokenAsync();
  if (aadtoken) {
    console.log(aadtoken);
    const tokenDiv = document.getElementById('token');
    tokenDiv.style.display = "none";
  } else {
    console.warn(`No token found. Please specify a token in the input box.`)
  }
}