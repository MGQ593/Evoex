/**
 * Genera iconos PNG para el add-in de Office
 * Usa solo módulos nativos de Node.js
 */

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

// Color azul Chevyplan
const BLUE = { r: 1, g: 84, b: 154 };  // #01549A
const WHITE = { r: 255, g: 255, b: 255 };

/**
 * Verifica si un punto está dentro de un rectángulo redondeado
 */
function isInRoundedRect(x, y, left, top, right, bottom, cornerRadius) {
  // Si está en el área central (sin esquinas), está dentro
  if (x >= left + cornerRadius && x <= right - cornerRadius && y >= top && y <= bottom) {
    return true;
  }
  if (y >= top + cornerRadius && y <= bottom - cornerRadius && x >= left && x <= right) {
    return true;
  }

  // Verificar las 4 esquinas redondeadas
  const corners = [
    { cx: left + cornerRadius, cy: top + cornerRadius },      // Superior izquierda
    { cx: right - cornerRadius, cy: top + cornerRadius },     // Superior derecha
    { cx: left + cornerRadius, cy: bottom - cornerRadius },   // Inferior izquierda
    { cx: right - cornerRadius, cy: bottom - cornerRadius },  // Inferior derecha
  ];

  for (const corner of corners) {
    const dx = x - corner.cx;
    const dy = y - corner.cy;
    // Solo verificar si estamos en la zona de la esquina
    if ((x < left + cornerRadius || x > right - cornerRadius) &&
        (y < top + cornerRadius || y > bottom - cornerRadius)) {
      if (dx * dx + dy * dy <= cornerRadius * cornerRadius) {
        return true;
      }
    }
  }

  return false;
}

/**
 * Crea un PNG con un diseño cuadrado redondeado y chat bubble
 */
function createIconPNG(size) {
  const width = size;
  const height = size;
  const rawData = Buffer.alloc(width * height * 4);

  // Margen y radio de esquina proporcionales al tamaño
  const margin = Math.max(1, Math.floor(size * 0.05));
  const cornerRadius = Math.max(2, Math.floor(size * 0.15));

  // Límites del rectángulo principal
  const rectLeft = margin;
  const rectTop = margin;
  const rectRight = size - margin - 1;
  const rectBottom = size - margin - 1;

  // Dimensiones del chat bubble (centrado)
  const centerX = width / 2;
  const centerY = height / 2;
  const bubbleWidth = size * 0.5;
  const bubbleHeight = size * 0.35;
  const bubbleRadius = Math.max(1, Math.floor(size * 0.06));

  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const idx = (y * width + x) * 4;

      // Verificar si está dentro del rectángulo redondeado principal
      if (isInRoundedRect(x, y, rectLeft, rectTop, rectRight, rectBottom, cornerRadius)) {
        // Dentro del fondo azul

        // Límites del bubble blanco
        const bubbleLeft = centerX - bubbleWidth / 2;
        const bubbleRight = centerX + bubbleWidth / 2;
        const bubbleTop = centerY - bubbleHeight / 2 - size * 0.05;
        const bubbleBottom = centerY + bubbleHeight / 2 - size * 0.05;

        // Verificar si está en el bubble
        const inBubble = isInRoundedRect(x, y, bubbleLeft, bubbleTop, bubbleRight, bubbleBottom, bubbleRadius);

        // Pequeño triángulo/cola del chat en la parte inferior izquierda
        const tailCenterX = bubbleLeft + bubbleWidth * 0.25;
        const tailTop = bubbleBottom - 1;
        const tailBottom = bubbleBottom + size * 0.12;
        const tailWidth = size * 0.08;
        const inTail = y >= tailTop && y <= tailBottom &&
                       x >= tailCenterX - tailWidth / 2 - (y - tailTop) * 0.5 &&
                       x <= tailCenterX + tailWidth / 2 - (y - tailTop) * 0.3;

        if (inBubble || inTail) {
          // Bubble blanco
          rawData[idx] = WHITE.r;
          rawData[idx + 1] = WHITE.g;
          rawData[idx + 2] = WHITE.b;
          rawData[idx + 3] = 255;

          // Añadir 3 puntos azules dentro del bubble (indicador de IA/chat)
          if (inBubble) {
            const dotRadius = Math.max(1, size * 0.045);
            const dotY = (bubbleTop + bubbleBottom) / 2;
            const dotSpacing = bubbleWidth * 0.22;
            const dots = [
              { x: centerX - dotSpacing, y: dotY },
              { x: centerX, y: dotY },
              { x: centerX + dotSpacing, y: dotY }
            ];

            for (const dot of dots) {
              const dotDist = Math.sqrt((x - dot.x) ** 2 + (y - dot.y) ** 2);
              if (dotDist <= dotRadius) {
                rawData[idx] = BLUE.r;
                rawData[idx + 1] = BLUE.g;
                rawData[idx + 2] = BLUE.b;
                rawData[idx + 3] = 255;
              }
            }
          }
        } else {
          // Fondo azul
          rawData[idx] = BLUE.r;
          rawData[idx + 1] = BLUE.g;
          rawData[idx + 2] = BLUE.b;
          rawData[idx + 3] = 255;
        }
      } else {
        // Fuera del rectángulo - transparente
        rawData[idx] = 0;
        rawData[idx + 1] = 0;
        rawData[idx + 2] = 0;
        rawData[idx + 3] = 0;
      }
    }
  }

  return encodePNG(width, height, rawData);
}

/**
 * Codifica datos RGBA a formato PNG
 */
function encodePNG(width, height, rawData) {
  // Crear datos filtrados (filter byte 0 = None para cada scanline)
  const filteredData = Buffer.alloc(height * (width * 4 + 1));
  for (let y = 0; y < height; y++) {
    const filterOffset = y * (width * 4 + 1);
    const rawOffset = y * width * 4;
    filteredData[filterOffset] = 0; // Filter type: None
    rawData.copy(filteredData, filterOffset + 1, rawOffset, rawOffset + width * 4);
  }

  // Comprimir con zlib
  const compressed = zlib.deflateSync(filteredData, { level: 9 });

  // Construir PNG
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  // IHDR chunk
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;  // bit depth
  ihdr[9] = 6;  // color type (RGBA)
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace
  const ihdrChunk = createChunk('IHDR', ihdr);

  // IDAT chunk
  const idatChunk = createChunk('IDAT', compressed);

  // IEND chunk
  const iendChunk = createChunk('IEND', Buffer.alloc(0));

  return Buffer.concat([signature, ihdrChunk, idatChunk, iendChunk]);
}

/**
 * Crea un chunk PNG con CRC
 */
function createChunk(type, data) {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);

  const typeBuffer = Buffer.from(type, 'ascii');
  const crcData = Buffer.concat([typeBuffer, data]);
  const crc = crc32(crcData);

  const crcBuffer = Buffer.alloc(4);
  crcBuffer.writeUInt32BE(crc, 0);

  return Buffer.concat([length, typeBuffer, data, crcBuffer]);
}

/**
 * Calcula CRC-32 para PNG
 */
function crc32(data) {
  let crc = 0xffffffff;
  const table = makeCRCTable();

  for (let i = 0; i < data.length; i++) {
    crc = (crc >>> 8) ^ table[(crc ^ data[i]) & 0xff];
  }

  return (crc ^ 0xffffffff) >>> 0;
}

function makeCRCTable() {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[n] = c;
  }
  return table;
}

// Generar iconos
const sizes = [16, 32, 80, 128];
const assetsDir = path.join(__dirname, '..', 'assets');

// Crear directorio si no existe
if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir, { recursive: true });
}

console.log('Generando iconos para Excel AI Assistant...\n');

for (const size of sizes) {
  const filename = `icon-${size}.png`;
  const filepath = path.join(assetsDir, filename);

  const pngData = createIconPNG(size);
  fs.writeFileSync(filepath, pngData);

  console.log(`✓ Creado: ${filename} (${size}x${size}px)`);
}

// Limpiar placeholder
const placeholderPath = path.join(assetsDir, 'icon-16.png.placeholder');
if (fs.existsSync(placeholderPath)) {
  fs.unlinkSync(placeholderPath);
  console.log('\n✓ Eliminado: icon-16.png.placeholder');
}

console.log('\n¡Iconos generados exitosamente!');
console.log('Los iconos muestran un cuadrado redondeado azul (#01549A) con bubble de chat blanco');
