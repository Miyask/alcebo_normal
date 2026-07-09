import type { VercelRequest, VercelResponse } from '@vercel/node';
import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { html, filename } = req.body;
    if (!html) {
      return res.status(400).json({ error: 'No se recibió contenido HTML.' });
    }

    // 1. Extract variables from HTML spans
    const extractSpan = (className: string, defaultValue: string = ''): string => {
      const regex = new RegExp(`class="[^"]*\\b${className}\\b[^"]*"[^>]*>([\\s\\S]*?)<\\/span>`, 'i');
      const match = html.match(regex);
      if (match && match[1]) {
        return match[1].replace(/<[^>]+>/g, '').trim();
      }
      return defaultValue;
    };

    const refCode = extractSpan('ref-code-field', 'Ref-@@@@@@@@@@@');
    const clientName = extractSpan('client-name-field', '@@@@@@@@');
    const clientAddress = extractSpan('client-address-field', '@@@@@@@@');
    const postalCode = extractSpan('postal-code-field', '@@@@');
    const postalCodePrefix = extractSpan('postal-code-prefix-field', '@@');
    const attName = extractSpan('att-name-field', '@@@@@@@@');
    const day = extractSpan('day-field', '@@');
    const month = extractSpan('month-field', '@@@@@');
    const year = extractSpan('year-field', '@@');
    const plaga = extractSpan('plaga-field', '@@@@');
    const zonasAfectadas = extractSpan('zonas-afectadas-field', '@@@@@@@@ y @@@@@@@@');
    
    // Technical observations
    const introTecnica = extractSpan('transcription-field', '');
    const problemaPrincipal = extractSpan('problema-principal-field', '@@@@@@@@');
    const detalleAdicional = extractSpan('detalle-adicional-field', '@@@@@@@@');
    
    // Protection zones
    const zona1 = extractSpan('zona-1-field', '@@@@@@@@');
    const zona2 = extractSpan('zona-2-field', '@@@@@@@@');
    const zona3 = extractSpan('zona-3-field', '@@@@@@@@');
    
    // Prices
    const price1 = extractSpan('price-field-1', '@@@@@');
    const price2 = extractSpan('price-field-2', '@@@@@');
    const price3 = extractSpan('price-field-3', '@@@@@');
    
    // Tech & contact
    const tecnico = extractSpan('tecnico-field', '@@@@@@@@@@@');
    const telefono = extractSpan('telefono-field', '@@@@@@@@');

    // 2. Extract base64 images from HTML img tags
    const images: Record<string, string> = {};
    
    // Map by data-img-id first
    const imgRegex = /<img[^>]+src="data:image\/(jpeg|png);base64,([^"]+)"[^>]*data-img-id="([^"]+)"/gi;
    let match;
    while ((match = imgRegex.exec(html)) !== null) {
      images[match[3]] = match[2];
    }
    
    // Fallback: Map by order of appearance
    const imgRawRegex = /<img[^>]+src="data:image\/(jpeg|png);base64,([^"]+)"/gi;
    let idx = 0;
    let matchRaw;
    while ((matchRaw = imgRawRegex.exec(html)) !== null) {
      idx++;
      const base64 = matchRaw[2];
      const imgId = `img_template_${idx}`;
      if (!images[imgId]) {
        images[imgId] = base64;
      }
    }

    // 3. Load native Word template file Ppo-mail-2022.docx
    const templatePath = path.join(process.cwd(), 'Ppo-mail-2022.docx');
    if (!fs.existsSync(templatePath)) {
      return res.status(500).json({ error: 'La plantilla Ppo-mail-2022.docx no se encuentra en el servidor.' });
    }
    const templateContent = fs.readFileSync(templatePath);
    const zip = new PizZip(templateContent);
    let docXml = zip.file('word/document.xml').asText();

    // 4. Modify XML text placeholders in word/document.xml by index matching
    let atIdx = 0;
    docXml = docXml.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/gi, (match, content) => {
      if (content.includes('@@')) {
        atIdx++;
        switch (atIdx) {
          case 1: return `<w:t>${refCode}</w:t>`;
          case 2: return `<w:t>${clientName}</w:t>`;
          case 3: return `<w:t>${clientAddress}</w:t>`;
          case 4: return `<w:t>${postalCode}   Madrid</w:t>`;
          case 5: return `<w:t>${attName}</w:t>`;
          case 6: return `<w:t>${day}</w:t>`;
          case 7: return `<w:t>${month}</w:t>`;
          case 8: return `<w:t>${year}</w:t>`;
          case 9: return `<w:t>${clientAddress}</w:t>`;
          case 10: return `<w:t>${plaga}</w:t>`;
          case 11: return `<w:t>${zonasAfectadas}</w:t>`;
          case 12: {
            // Multi-paragraph technical description
            const lines = introTecnica.split('\n').filter((l: string) => l.trim().length > 0);
            if (lines.length > 0) {
              return lines.join('</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr><w:t>');
            }
            return `<w:t>se observó presencia activa de aves en la edificación</w:t>`;
          }
          case 13: return `<w:t>El problema principal ${problemaPrincipal}</w:t>`;
          case 14: return `<w:t>${detalleAdicional}</w:t>`;
          case 15: return `<w:t>${zona1}</w:t>`;
          case 16: return `<w:t>${zona2}</w:t>`;
          case 17: return `<w:t>${zona3}</w:t>`;
          case 18: return `<w:t>${telefono}</w:t>`;
          case 19: return `<w:t>280${postalCodePrefix}</w:t>`;
          case 20: return `<w:t>${refCode}</w:t>`;
          case 21: return `<w:t>................ ${price1}</w:t>`;
          case 22: return `<w:t>${price3}</w:t>`;
          case 23: return `<w:t>........ ${price2}</w:t>`;
          case 24: return `<w:t>${tecnico}</w:t>`;
          case 25: return `<w:t>${clientAddress}</w:t>`;
        }
      }
      return match;
    });

    // 5. Overwrite/replace images in ZIP package
    if (images['img_template_2']) {
      zip.file('word/media/image2.jpeg', Buffer.from(images['img_template_2'], 'base64'), { binary: true });
    }
    if (images['img_template_3']) {
      zip.file('word/media/image3.jpeg', Buffer.from(images['img_template_3'], 'base64'), { binary: true });
    }
    if (images['img_template_4']) {
      zip.file('word/media/image4.jpeg', Buffer.from(images['img_template_4'], 'base64'), { binary: true });
    }

    // 6. Delete unused images from word/document.xml if they were removed in the editor
    if (!html.includes('data-img-id="img_template_2"') && !html.includes('Foto_Inspeccion_1.jpg')) {
      const pRegex = /<w:p[^>]*>[\s\S]*?r:embed="rId11"[\s\S]*?<\/w:p>/g;
      docXml = docXml.replace(pRegex, '');
    }
    if (!html.includes('data-img-id="img_template_3"') && !html.includes('Foto_Inspeccion_2.jpg')) {
      const pRegex = /<w:p[^>]*>[\s\S]*?r:embed="rId12"[\s\S]*?<\/w:p>/g;
      docXml = docXml.replace(pRegex, '');
    }
    if (!html.includes('data-img-id="img_template_4"') && !html.includes('Propuesta_Tecnica.jpg')) {
      const pRegex = /<w:p[^>]*>[\s\S]*?r:embed="rId14"[\s\S]*?<\/w:p>/g;
      docXml = docXml.replace(pRegex, '');
    }

    // 7. Write modified XML back into zip
    zip.file('word/document.xml', docXml);

    // 8. Generate DOCX file buffer and send response
    const fileBuffer = zip.generate({ type: 'nodebuffer' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename || 'Presupuesto'}.docx"`);
    return res.status(200).send(fileBuffer);

  } catch (error: any) {
    console.error('Error al generar el documento Word desde la plantilla:', error);
    return res.status(500).json({ error: 'Error al rellenar la plantilla Word.', details: error.message });
  }
}
