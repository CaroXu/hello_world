const express = require('express');
const path = require("path");

const readXlsxFile = require('read-excel-file/node');
const {readSheetNames} = require('read-excel-file/node');

const app = express()
const port = 3000

const clientExcelFile = 'data/Site_Audit.xlsx';
const verificationText = 'Google';

const posClientName = 0;
const posClientURL = 1;
const maxTimeToWaitForSite = 10;

app.set("view engine","pug");
app.set("views", path.join(__dirname,"views"))

function client(url, status){
  this.url = url;
  this.status = status;
}

app.get("/", (req, res) => {
  readClientsFromExcel(clientExcelFile).then((serverInfo) => {
    // serverInfo contains all known servers and their clients
    outputClientSiteStatus(serverInfo,res);
  });
});

app.listen(port, () => {
  console.log(`CheckClientStatus app listening on port ${port}`)
});

async function readClientsFromExcel (filename) {
  const sheetNames = await readSheetNames(filename);

  const serverInfo = {
    summary: [],
    servers: {}
  };

  for (const sheetName of sheetNames) {
    if (sheetName === 'Summary') {
      serverInfo.summary = await readXlsxFile(filename, { sheet: sheetName });
    } else {
      serverInfo.servers[sheetName] = await readXlsxFile(filename, { sheet: sheetName });
    }
  }

  return serverInfo;
}

async function outputClientSiteStatus (serverInfo,resp) {
  let data = [];
  let clientUrls = [];
  const startTime = new Date();
  for (const serverName in serverInfo.servers) {
    serverInfo.servers[serverName].forEach(element => {
      const clientName = element[posClientName];
      const clientURL = element[posClientURL];
      // Only process client site
      if ((clientURL !== null) && (clientURL !== 'Site URL')) {
        clientUrls.push(clientURL);
      }
    });
  }

  const promiseClients = clientUrls.map(x => { return checkClient(x) });
  Promise.allSettled(promiseClients)
    .then((values) => {
    // all promises have been evaluated
    // determine which clients (if any) have issues
      const endTime = new Date();
      let errorClientNum = 0;
      for (const x in values) {
        let clientstatus = false;
        const clientURL = clientUrls[x];
        if (values[x].status === 'fulfilled') {
          clientstatus = true;
        }else{
          errorClientNum ++;
        }
        data.push(new client(clientURL, clientstatus));
      }

      // export the data and render the front end page
      try {
        const clientNum = data.length;
        console.log(`finished getting status of all sites, ran from ${startTime} to ${endTime}`);
        resp.render("CheckAllClientStatus",{ title:'Sites Status',
        data: data, errorClientNum: errorClientNum, clientNum: clientNum});
      } catch (err) {
        console.error(err);
      }
    })
}

// test a single client, resolve if ok else reject
function checkClient (clientURL) {
  console.log('checking: '+clientURL );
  return fetchWithTimeout(clientURL)
    .then(res => {
    // verify the response was ok, eg HTTP 200
      if (!res.ok) { throw new Error(`HTTP error: ${res.status}`); }
      return res.text()
    })
    .then(text => {
      // verify the response contains the text we want
      if (!text.includes(verificationText)) { throw new Error(`Required text was not found`); }
      // client was verified ok
      return;
    });
// do not catch errors, allow them to bubble up
}

// borrowed from https://dmitripavlutin.com/timeout-fetch-request/
// wait a max of "maxTimeToWaitForSite" seconds for a network response
async function fetchWithTimeout (resource, options = {}) {
  const { timeout = maxTimeToWaitForSite * 1000 } = options;

  const controller = new AbortController();
  const id = setTimeout(() => controller.abort(), timeout);
  const response = await fetch(resource, {
    ...options,
    signal: controller.signal
  });
  clearTimeout(id);
  return response;
}