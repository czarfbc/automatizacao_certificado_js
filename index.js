import XLSX from 'xlsx';
import puppeteer from 'puppeteer';
import fs from 'fs/promises';

const workbook = XLSX.readFile('Planilha sem título.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const jsonData = XLSX.utils.sheet_to_json(worksheet);

(async () => {
  const browser = await puppeteer.launch({ headless: false });

  for (const element of jsonData) {
    const page = await browser.newPage();

    try {
      const htmlPath = 'index.html';
      const htmlContent = await fs.readFile(htmlPath, 'utf-8');

      const content = htmlContent
        .replace('{{clientes}}', element.clientes)
        .replace('{{valor}}', element.valor)
        .replace('{{cpf}}', element.cpf);

      await page.setContent(content);

      const pdfPath = `./certificados/${element.clientes}.pdf`;
      await page.pdf({ path: pdfPath, width: '800px', height: '800' });

      console.log(`PDF gerado para o cliente ${element.clientes}: ${pdfPath}`);
    } catch (error) {
      console.error(
        `Erro ao processar o cliente ${element.clientes}: ${error.message}`
      );
    } finally {
      await page.close();
    }
  }

  await browser.close();
  console.log('PDFs gerados com sucesso!');
})();
