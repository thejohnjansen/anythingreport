'use strict';
/**
 * Generates logo.png (256×256) and electron/icon.ico from logo.svg.
 * Run once after changing logo.svg:  node generate-icon.js
 */
const path = require('path');
const fs   = require('fs');
const sharp    = require('sharp');
const pngToIco = require('png-to-ico');

const SVG_PATH = path.join(__dirname, 'logo.svg');
const PNG_PATH = path.join(__dirname, 'logo.png');
const ICO_PATH = path.join(__dirname, 'electron', 'icon.ico');

(async () => {
    console.log('Rendering SVG → PNG (256×256)…');
    await sharp(SVG_PATH)
        .resize(256, 256)
        .png()
        .toFile(PNG_PATH);
    console.log(`  ✓ ${PNG_PATH}`);

    console.log('Converting PNG → ICO…');
    const icoBuffer = await (pngToIco.default || pngToIco)([PNG_PATH]);
    fs.writeFileSync(ICO_PATH, icoBuffer);
    console.log(`  ✓ ${ICO_PATH}`);

    console.log('Done.');
})();
