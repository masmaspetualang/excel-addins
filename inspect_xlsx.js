const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

// Since we don't have adm-zip, we can execute a PowerShell command to unzip it!
// Or we can just use PowerShell's Expand-Archive!
console.log("To be executed via terminal command.");
