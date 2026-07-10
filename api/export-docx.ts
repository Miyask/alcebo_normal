import type { VercelRequest, VercelResponse } from '@vercel/node';
import { WATERMARK_BASE64 } from '../src/data/watermarkBase64';

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { html, filename } = req.body;
    if (!html) {
      return res.status(400).json({ error: 'No se recibió contenido HTML.' });
    }

    // Dynamic import to catch load time native module errors in Vercel environment
    let htmlToDocx;
    try {
      htmlToDocx = (await import('html-to-docx')).default;
    } catch (importErr: any) {
      console.error('Failed to import html-to-docx:', importErr);
      return res.status(500).json({ 
        error: 'Error al cargar el compilador html-to-docx en el servidor.',
        details: importErr.message,
        stack: importErr.stack
      });
    }

    // Add high-fidelity styled page wrapper to the HTML for conversion
    const fullHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <style>
          body {
            font-family: 'Calibri', 'Arial', sans-serif;
            font-size: 11pt;
            line-height: 1.5;
          }
          p {
            margin-bottom: 10pt;
            text-align: justify;
          }
          h1, h2, h3, h4 {
            color: #009FE3;
            font-family: 'Calibri', 'Arial', sans-serif;
            margin-top: 18pt;
            margin-bottom: 6pt;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 12pt;
            margin-bottom: 12pt;
          }
          th, td {
            border: 1px solid #bec8d2;
            padding: 6pt;
            font-size: 10pt;
          }
          th {
            background-color: #009FE3;
            color: #ffffff;
            font-weight: bold;
          }
          /* Custom styles for A4 layout elements */
          .cover-page-wrapper {
            text-align: center;
          }
          .image-wrapper {
            text-align: center;
            margin: 20px 0;
          }
          .document-image {
            max-width: 100%;
            height: auto;
            border: 1px solid #bec8d2;
            border-radius: 8px;
          }
        </style>
      </head>
      <body>
        ${html}
      </body>
      </html>
    `;

    const headerHtml = `
      <div style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; z-index: -1000; opacity: 0.08; text-align: center;">
        <img src="data:image/jpeg;base64,${WATERMARK_BASE64}" style="width: 550px; margin-top: 250px;" />
      </div>
    `;

    let fileBuffer;
    try {
      fileBuffer = await htmlToDocx(fullHtml, headerHtml, {
        table: { row: { cantSplit: true } },
        footer: true,
        header: true,
        pageNumber: true,
      });
    } catch (convErr: any) {
      console.error('html-to-docx conversion crashed:', convErr);
      return res.status(500).json({
        error: 'El compilador de Word ha fallado durante la conversión.',
        details: convErr.message,
        stack: convErr.stack
      });
    }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename || 'Presupuesto'}.docx"`);
    return res.status(200).send(fileBuffer);

  } catch (error: any) {
    console.error('General error in export-docx:', error);
    return res.status(500).json({ error: 'Error general de exportación.', details: error.message, stack: error.stack });
  }
}
