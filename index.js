import XLSX from 'xlsx';
import puppeteer from 'puppeteer';
import fs from 'fs/promises';

const workbook = XLSX.readFile('./planilha/Pasta.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const jsonData = XLSX.utils.sheet_to_json(worksheet);

console.log(jsonData);

function removerAcentos(str) {
  return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

(async () => {
  const browser = await puppeteer.launch();

  for (const element of jsonData) {
    const page = await browser.newPage();

    try {
      const htmlPath = 'index.html';
      const htmlContent = await fs.readFile(htmlPath, 'utf-8');

      const nomeSemAcento = removerAcentos(element.Colaborador);

      const content = htmlContent
        .replace('{{ColaboradorTitle}}', nomeSemAcento)
        .replace('{{CPF}}', element.CPF)
        .replace('{{TopicoAssunto}}', element.TopicoAssunto)
        .replace('{{Duracao}}', element.Duracao)
        .replace('{{Data}}', element.Data)
        .replace('{{ColaboradorAss}}', element.Colaborador);

      await page.setContent(content);

      const pdfPath = `./certificados/${element.Colaborador}.pdf`;
      await page.pdf({
        path: pdfPath,
        printBackground: true,
        width: '960px',
        height: '718px',
        pageRanges: '1-2',
      });

      console.log(
        `PDF gerado para o cliente ${element.Colaborador}: ${pdfPath}`
      );
    } catch (error) {
      console.error(
        `Erro ao processar o cliente ${element.Colaborador}: ${error.message}`
      );
    } finally {
      await page.close();
    }
  }

  await browser.close();
  console.log('PDFs gerados com sucesso!');
})();
