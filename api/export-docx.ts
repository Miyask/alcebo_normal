import type { VercelRequest, VercelResponse } from '@vercel/node';
import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { html, filename, variables } = req.body;
    if (!html) {
      return res.status(400).json({ error: 'No se recibió contenido HTML.' });
    }

    // 1. Retrieve variables from body or parse them from HTML (fallback)
    const extractContent = (className: string, defaultValue: string = ''): string => {
      if (variables && variables[className]) {
        return variables[className];
      }
      const regex = new RegExp(`class="[^"]*\\b${className}\\b[^"]*"[^>]*>([\\s\\S]*?)<\\/(?:span|div|p)>`, 'i');
      const match = html.match(regex);
      if (match && match[1]) {
        return match[1].replace(/<[^>]+>/g, '').trim();
      }
      return defaultValue;
    };

    const refCode = variables?.refCode || extractContent('ref-code-field', 'Ref-@@@@@@@@@@@');
    const clientName = variables?.clientName || extractContent('client-name-field', '@@@@@@@@');
    const clientAddress = variables?.clientAddress || extractContent('client-address-field', '@@@@@@@@');
    const postalCode = variables?.postalCode || extractContent('postal-code-field', '@@@@');
    const postalCodePrefix = variables?.postalCodePrefix || extractContent('postal-code-prefix-field', '@@');
    const attName = variables?.attName || extractContent('att-name-field', '@@@@@@@@');
    const day = variables?.day || extractContent('day-field', '@@');
    const month = variables?.month || extractContent('month-field', '@@@@@');
    const year = variables?.year || extractContent('year-field', '@@');
    const plaga = variables?.plaga || extractContent('plaga-field', '@@@@');
    const zonasAfectadas = variables?.zonasAfectadas || extractContent('zonas-afectadas-field', '@@@@@@@@ y @@@@@@@@');
    
    // Technical observations
    const introTecnica = variables?.introTecnica || extractContent('transcription-field', '');
    const problemaPrincipal = variables?.problemaPrincipal || extractContent('problema-principal-field', '@@@@@@@@');
    const detalleAdicional = variables?.detalleAdicional || extractContent('detalle-adicional-field', '@@@@@@@@');
    
    // Protection zones
    const zona1 = variables?.zona1 || extractContent('zona-1-field', '@@@@@@@@');
    const zona2 = variables?.zona2 || extractContent('zona-2-field', '@@@@@@@@');
    const zona3 = variables?.zona3 || extractContent('zona-3-field', '@@@@@@@@');
    
    // Prices
    const price1 = variables?.price1 || extractContent('price-field-1', '@@@@@');
    const price2 = variables?.price2 || extractContent('price-field-2', '@@@@@');
    const price3 = variables?.price3 || extractContent('price-field-3', '@@@@@');
    
    // Tech & contact
    const tecnico = variables?.tecnico || extractContent('tecnico-field', '@@@@@@@@@@@');
    const telefono = variables?.telefono || extractContent('telefono-field', '@@@@@@@@');

    // Bird plaga description block
    const plagaDescription = variables?.plagaDescription || extractContent('des-plaga-block', '');

    // Active systems
    const activeSystems = variables?.activeSystems || (html.includes('RED NETWORK ANTI-PALOMAS') ? ['Red'] : (html.includes('VARILLAS AVIPOINT') ? ['Varillas'] : []));

    // 2. Extract base64 custom visit images from HTML img tags (img_template_2 and img_template_3 only)
    const images: Record<string, string> = {};
    const imgRegex = /<img[^>]+src="data:image\/(jpeg|png);base64,([^"]+)"[^>]*data-img-id="([^"]+)"/gi;
    let match;
    while ((match = imgRegex.exec(html)) !== null) {
      if (match[3].startsWith('img_template_')) {
        images[match[3]] = match[2];
      }
    }
    
    // Fallback: Map by order of appearance
    const imgRawRegex = /<img[^>]+src="data:image\/(jpeg|png);base64,([^"]+)"/gi;
    let idx = 0;
    let matchRaw;
    while ((matchRaw = imgRawRegex.exec(html)) !== null) {
      idx++;
      const base64 = matchRaw[2];
      // Skip logo and only map visit photos
      if (idx > 1) {
        const imgId = `img_template_${idx}`;
        if (!images[imgId]) {
          images[imgId] = base64;
        }
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
          case 10: {
            const pluralMap: Record<string, string> = {
              'Palomas': 'palomas',
              'Gorriones': 'gorriones',
              'Cigüeñas': 'cigüeñas',
              'Gaviotas': 'gaviotas',
              'Cotorras': 'cotorras',
              'Golondrinas': 'golondrinas',
              'Urracas': 'urracas'
            };
            const pluralBird = pluralMap[plaga] || plaga.toLowerCase();
            
            const searchStr = '<w:t xml:space="preserve">palomas en </w:t>';
            const nextPalomas = docXml.indexOf(searchStr);
            if (nextPalomas !== -1) {
              docXml = docXml.substring(0, nextPalomas) + '<w:t xml:space="preserve"> en </w:t>' + docXml.substring(nextPalomas + searchStr.length);
            }
            return `<w:t>${pluralBird}</w:t>`;
          }
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
          case 19: return `<w:t>${postalCodePrefix}</w:t>`;
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

    // 5. Dynamic Bird Description injection under Section 1 in XML
    if (plagaDescription) {
      const birdAnchorRegex = /<w:p[^>]*>[\s\S]*?aprovechar los desechos animales[\s\S]*?<\/w:p>/i;
      const lines = plagaDescription.split('\n').filter(l => l.trim().length > 0);
      const xmlParagraphs = lines.map(line => `
        <w:p><w:pPr><w:jc w:val="both"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr></w:pPr>
          <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:i/></w:rPr>
            <w:t>${line}</w:t>
          </w:r>
        </w:p>
      `).join('');
      
      docXml = docXml.replace(birdAnchorRegex, (match) => match + xmlParagraphs);
    }

    // 6. Dynamic System Specifications injection and removal in XML
    // Remove RED description and diagram if not proposed
    if (!activeSystems.includes('Red')) {
      const pRedTitle = /<w:p[^>]*>[\s\S]*?r:id="rId11"[\s\S]*?<\/w:p>/g;
      docXml = docXml.replace(pRedTitle, '');
      const pRedBullets = [
        /<w:p[^>]*>[\s\S]*?Base de polietileno trenzado[\s\S]*?<\/w:p>/gi,
        /<w:p[^>]*>[\s\S]*?Fijación de la red sobre cable[\s\S]*?<\/w:p>/gi,
        /<w:p[^>]*>[\s\S]*?Cada hebra se forma por 3 filamentos[\s\S]*?<\/w:p>/gi,
        /<w:p[^>]*>[\s\S]*?El diámetro del rombo[\s\S]*?<\/w:p>/gi,
      ];
      pRedBullets.forEach(re => { docXml = docXml.replace(re, ''); });
    }

    // Remove VARILLAS description and diagram if not proposed
    if (!activeSystems.includes('Varillas')) {
      const pVarillasTitle = /<w:p[^>]*>[\s\S]*?r:id="rId12"[\s\S]*?<\/w:p>/g;
      docXml = docXml.replace(pVarillasTitle, '');
      const pVarillasBullets = [
        /<w:p[^>]*>[\s\S]*?Alambre de acero inoxidable[\s\S]*?<\/w:p>/gi,
        /<w:p[^>]*>[\s\S]*?Punta roma de baja[\s\S]*?<\/w:p>/gi,
        /<w:p[^>]*>[\s\S]*?Fijación con adhesivo sellador[\s\S]*?<\/w:p>/gi,
      ];
      pVarillasBullets.forEach(re => { docXml = docXml.replace(re, ''); });
    }

    // Append Eléctrico or Capturas if proposed
    const electricoXml = `
      <w:p><w:pPr><w:jc w:val="both"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr></w:pPr>
        <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:b/></w:rPr>
          <w:t>SISTEMA ELECTROESTÁTICO DISUASORIO (ELÉCTRICO): </w:t>
        </w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr>
          <w:t>sus características son las siguientes:</w:t>
        </w:r>
      </w:p>
      <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="17"/></w:numPr><w:jc w:val="both"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr></w:pPr>
        <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr>
          <w:t>Solución de alta discreción visual, ideal para edificios catalogados o zonas de alto valor estético.</w:t>
        </w:r>
      </w:p>
      <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="17"/></w:numPr><w:jc w:val="both"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr></w:pPr>
        <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr>
          <w:t>Emisión de impulsos electroestáticos de baja frecuencia y baja intensidad, completamente inocuos para las aves pero altamente disuasorios.</w:t>
        </w:r>
      </w:p>
      <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="17"/></w:numPr><w:jc w:val="both"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr></w:pPr>
        <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr>
          <w:t>Línea perimetral de conductores de acero inoxidable fijados sobre aisladores de policarbonato estabilizado.</w:t>
        </w:r>
      </w:p>
    `;

    const capturasXml = `
      <w:p><w:pPr><w:jc w:val="both"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr></w:pPr>
        <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:b/></w:rPr>
          <w:t>PLAN DE CAPTURAS SELECTIVAS: </w:t>
        </w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr>
          <w:t>sus características son las siguientes:</w:t>
        </w:r>
      </w:p>
      <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="17"/></w:numPr><w:jc w:val="both"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr></w:pPr>
        <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr>
          <w:t>Instalación de jaulas trampa homologadas dotadas de comederos, bebederos y sombreado para garantizar el bienestar animal.</w:t>
        </w:r>
      </w:p>
      <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="17"/></w:numPr><w:jc w:val="both"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr></w:pPr>
        <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr>
          <w:t>Revisiones y mantenimiento periódico por técnicos autorizados para control de capturas, retirada selectiva y cebado.</w:t>
        </w:r>
      </w:p>
      <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="17"/></w:numPr><w:jc w:val="both"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr></w:pPr>
        <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/></w:rPr>
          <w:t>Retirada y traslado humanitario de los ejemplares de acuerdo con la legislación autonómica de protección y sanidad animal.</w:t>
        </w:r>
      </w:p>
    `;

    if (activeSystems.includes('Eléctrico') || activeSystems.includes('Capturas')) {
      const anchorRegex = /<w:p[^>]*>[\s\S]*?A continuación detallamos las características de los sistemas elegidos[\s\S]*?<\/w:p>/i;
      docXml = docXml.replace(anchorRegex, (match) => {
        let extraXml = '';
        if (activeSystems.includes('Eléctrico')) {
          extraXml += electricoXml;
        }
        if (activeSystems.includes('Capturas')) {
          extraXml += capturasXml;
        }
        return match + extraXml;
      });
    }

    // 7. Inject custom visit photos dynamically inside XML by replacing "Foto Muestra" tags
    let relsXml = zip.file('word/_rels/document.xml.rels').asText();
    if (images['img_template_2']) { // Visit photo 1
      relsXml = relsXml.replace('</Relationships>', '<Relationship Id="rId25" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image5.jpeg"/></Relationships>');
      zip.file('word/media/image5.jpeg', Buffer.from(images['img_template_2'], 'base64'), { binary: true });
    }
    if (images['img_template_3']) { // Visit photo 2
      relsXml = relsXml.replace('</Relationships>', '<Relationship Id="rId26" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image6.jpeg"/></Relationships>');
      zip.file('word/media/image6.jpeg', Buffer.from(images['img_template_3'], 'base64'), { binary: true });
    }
    zip.file('word/_rels/document.xml.rels', relsXml);

    let photoIdx = 0;
    docXml = docXml.replace(/<w:p[^>]*>([\s\S]*?<w:t>Foto Muestra<\/w:t>[\s\S]*?)<\/w:p>/gi, (match) => {
      photoIdx++;
      if (photoIdx === 1 && images['img_template_2']) {
        return `
          <w:p>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r>
              <w:pict>
                <v:shape id="VisitPhoto1" style="width:360pt;height:240pt;" type="#_x0000_t75">
                  <v:imagedata r:id="rId25" o:title="Foto Inspeccion 1"/>
                </v:shape>
              </w:pict>
            </w:r>
          </w:p>
        `;
      }
      if (photoIdx === 2 && images['img_template_3']) {
        return `
          <w:p>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r>
              <w:pict>
                <v:shape id="VisitPhoto2" style="width:360pt;height:240pt;" type="#_x0000_t75">
                  <v:imagedata r:id="rId26" o:title="Foto Inspeccion 2"/>
                </v:shape>
              </w:pict>
            </w:r>
          </w:p>
        `;
      }
      return ''; // Delete the "Foto Muestra" placeholder text if no photo uploaded
    });

    // 8. Write modified XML back into zip
    zip.file('word/document.xml', docXml);

    // 9. Generate DOCX file buffer and send response
    const fileBuffer = zip.generate({ type: 'nodebuffer' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename || 'Presupuesto'}.docx"`);
    return res.status(200).send(fileBuffer);

  } catch (error: any) {
    console.error('Error al generar el documento Word desde la plantilla:', error);
    return res.status(500).json({ error: 'Error al rellenar la plantilla Word.', details: error.message });
  }
}
