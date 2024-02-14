//import XLSX from 'xlsx';
import puppeteer from 'puppeteer';
import fs from 'fs/promises';
import mammoth from 'mammoth';

// const workbook = XLSX.readFile('./planilha/Pasta.xlsx');
// const worksheet = workbook.Sheets[workbook.SheetNames[0]];
// const jsonData = XLSX.utils.sheet_to_json(worksheet);

// console.log(jsonData);

let docxPath = './planilha/dd.docx';
let jsonData = [];

mammoth
  .extractRawText({ path: docxPath })
  .then(function (result) {
    let text = result.value;
    let lines = text.split('\n');

    for (let line of lines) {
      let trimmedLine = line.trim();

      // Verificar se a linha parece ser um nome de colaborador
      if (/^[^\d]+$/g.test(trimmedLine)) {
        jsonData.push({ Colaborador: trimmedLine });
      }

      // Verificar se a linha parece ser um CPF
      if (
        /^\d{3}\.\d{3}\.\d{3}-\d{2}$/g.test(trimmedLine) ||
        /^\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}$/g.test(trimmedLine)
      ) {
        if (jsonData.length > 0) {
          jsonData[jsonData.length - 1]['CPF'] = trimmedLine;
        }
      }
    }

    console.log(jsonData);
  })
  .catch(function (err) {
    console.log(err);
  });

function removerAcentos(str) {
  return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

let numbersPDF = 0;

(async () => {
  const browser = await puppeteer.launch();

  for (const element of jsonData) {
    const page = await browser.newPage();

    try {
      const htmlPath = 'index.html';
      const htmlContent = await fs.readFile(htmlPath, 'utf-8');

      const nomeSemAcento = removerAcentos(element.Colaborador);
      const nomeMaiusculoSemAcento = nomeSemAcento.toUpperCase();

      const content = htmlContent
        .replace('{{ColaboradorTitle}}', nomeMaiusculoSemAcento)
        .replace('{{CPF}}', element.CPF)
        .replace('{{ColaboradorAss}}', element.Colaborador);

      await page.setContent(content);

      const namePDF = element.Colaborador.split(' ').join('_');
      const cpfPDF = element.CPF.split(/[./-]/).join('');

      const pdfPath = `./certificados/${namePDF}_${cpfPDF}.pdf`;
      await page.pdf({
        path: pdfPath,
        printBackground: true,
        width: '960px',
        height: '720px',
        pageRanges: '1-2',
      });

      console.log(
        '\x1b[38;2;0;255;0m%s\x1b[0m',
        `PDF gerado para o cliente ${element.Colaborador}: ${pdfPath}`
      );

      numbersPDF++;
    } catch (error) {
      console.error(
        '\x1b[38;2;255;0;0m%s\x1b[0m',
        `Erro ao processar o cliente ${element.Colaborador}: ${error.message}`
      );
    } finally {
      await page.close();
    }
  }

  await browser.close();
  console.log(
    '\x1b[38;2;0;255;255m%s\x1b[0m',
    `${numbersPDF} PDFs gerados com sucesso! `
  );
})();
