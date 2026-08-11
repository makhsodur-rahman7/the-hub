import fs from 'node:fs/promises';
import path from 'node:path';

async function downloadFavicon() {
  console.log('Fetching https://thehubonverulam.com/ ...');
  
  let faviconUrl = 'https://thehubonverulam.com/favicon.ico';
  let ext = '.ico';
  
  try {
    const response = await fetch('https://thehubonverulam.com/');
    const html = await response.text();
    
    // Look for link rel="icon" or apple-touch-icon
    const iconRegex = /<link[^>]*rel=["'](?:shortcut )?(?:apple-touch-)?icon["'][^>]*href=["']([^"']+)["'][^>]*>/i;
    const match = html.match(iconRegex);
    
    if (match && match[1]) {
      let href = match[1];
      if (href.startsWith('/')) {
        href = 'https://thehubonverulam.com' + href;
      }
      faviconUrl = href;
      ext = path.extname(new URL(faviconUrl).pathname) || '.ico';
      console.log(`Found icon link in HTML: ${faviconUrl}`);
    }
    
    console.log(`Downloading favicon from ${faviconUrl} ...`);
    const iconRes = await fetch(faviconUrl);
    
    if (!iconRes.ok) {
      throw new Error(`Failed to download icon: ${iconRes.status} ${iconRes.statusText}`);
    }
    
    const buffer = await iconRes.arrayBuffer();
    
    // Ensure public directory exists
    try {
      await fs.access('./public');
    } catch {
      await fs.mkdir('./public');
    }
    
    const destPath = `./public/favicon${ext}`;
    await fs.writeFile(destPath, Buffer.from(buffer));
    console.log(`Saved favicon to ${destPath}`);
    
    // Update index.html
    let indexHtml = await fs.readFile('./index.html', 'utf8');
    
    // Check if link rel="icon" already exists
    if (!indexHtml.includes('rel="icon"')) {
      const headEndIdx = indexHtml.indexOf('</head>');
      if (headEndIdx !== -1) {
        const linkTag = `    <link rel="icon" href="/favicon${ext}" />\n`;
        indexHtml = indexHtml.slice(0, headEndIdx) + linkTag + indexHtml.slice(headEndIdx);
        await fs.writeFile('./index.html', indexHtml);
        console.log('Updated index.html to include the new favicon.');
      }
    } else {
      console.log('index.html already contains a favicon link, please update it manually if needed.');
    }
    
  } catch (error) {
    console.error('Error:', error);
  }
}

downloadFavicon();
