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
      const nomeMaiusculo = nomeSemAcento.toUpperCase();

      const dataExcel = new Date(
        (element.Data - 1) * 24 * 3600 * 1000 + new Date('1900-01-01').getTime()
      );

      const dia = dataExcel.getDate().toString().padStart(2, '0');
      let mes = (dataExcel.getMonth() + 1).toString().padStart(2, '0');
      const ano = dataExcel.getFullYear();

      switch (mes) {
        case '01':
          mes = 'Janeiro';
          break;
        case '02':
          mes = 'Fevereiro';
          break;
        case '03':
          mes = 'Março';
          break;
        case '04':
          mes = 'Abril';
          break;
        case '05':
          mes = 'Maio';
          break;
        case '06':
          mes = 'Junho';
          break;
        case '07':
          mes = 'Julho';
          break;
        case '08':
          mes = 'Agosto';
          break;
        case '09':
          mes = 'Setembro';
          break;
        case '10':
          mes = 'Outubro';
          break;
        case '11':
          mes = 'Novembro';
          break;
        case '12':
          mes = 'Dezembro';
          break;
        default:
          break;
      }

      const dataFormatada = `${dia} de ${mes} de ${ano}`;

      const content = htmlContent
        .replace('{{ColaboradorTitle}}', nomeMaiusculo)
        .replace('{{CPF}}', element.CPF)
        .replace('{{TopicoAssunto}}', element.TopicoAssunto)
        .replace('{{Duracao}}', element.Duracao)
        .replace('{{Data}}', dataFormatada)
        .replace('{{ColaboradorAss}}', element.Colaborador);

      await page.setContent(content);

      const pdfPath = `./certificados/${element.Colaborador}.pdf`;
      await page.pdf({
        path: pdfPath,
        printBackground: true,
        width: '960px',
        height: '720px',
        // format: 'A4',
        // landscape: true,
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
