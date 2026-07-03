const fs = require('fs');
const https = require('https');
const http = require('http');
const path = require('path');

function download(url, dest) {
  return new Promise((resolve, reject) => {
    const protocol = url.startsWith('https') ? https : http;
    protocol.get(url, (response) => {
      if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
        // Resolve absolute or relative redirect URL
        const redirectUrl = new URL(response.headers.location, url).href;
        download(redirectUrl, dest).then(resolve).catch(reject);
        return;
      }
      
      if (response.statusCode !== 200) {
        reject(new Error(`Failed to download ${url}: status code ${response.statusCode}`));
        return;
      }
      
      const file = fs.createWriteStream(dest);
      response.pipe(file);
      file.on('finish', () => {
        file.close();
        console.log(`Successfully downloaded ${url} to ${dest}`);
        resolve();
      });
    }).on('error', (err) => {
      reject(err);
    });
  });
}

const libDir = path.join(__dirname, '..', 'public', 'js', 'lib');
if (!fs.existsSync(libDir)) {
  fs.mkdirSync(libDir, { recursive: true });
}

const tasks = [
  {
    url: 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/dist/umd/supabase.js',
    dest: path.join(libDir, 'supabase.js')
  },
  {
    url: 'https://unpkg.com/lucide@latest/dist/umd/lucide.js',
    dest: path.join(libDir, 'lucide.js')
  },
  {
    url: 'https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.all.min.js',
    dest: path.join(libDir, 'sweetalert2.js')
  }
];

Promise.all(tasks.map(t => download(t.url, t.dest)))
  .then(() => console.log('All downloads completed!'))
  .catch(err => {
    console.error('Error during downloading:', err);
    process.exit(1);
  });
