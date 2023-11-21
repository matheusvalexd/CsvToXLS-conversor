const express = require('express');
const csvtojson = require('csvtojson');
const Excel = require('excel4node');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

const app = express();

app.use(express.json());

let convertedFilePath = null; // Caminho do arquivo convertido

app.post('/converter', async (req, res) => {
  try {
    const { csvUrl, fileName } = req.body;

    if (!csvUrl) {
      return res.status(400).json({ error: 'Envie o link do arquivo CSV no corpo da requisição.' });
    }

    const response = await axios.get(csvUrl);
    const csvData = response.data;

    const jsonArray = await csvtojson().fromString(csvData);

    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('Sheet 1');

    const style = wb.createStyle({
      alignment: {
        horizontal: 'center',
      },
    });

    // Verifica se há dados no CSV
    if (jsonArray.length > 0) {
      const headers = Object.keys(jsonArray[0]); // Obtém os cabeçalhos do CSV
      const columnMaxLength = {};

      // Encontra o tamanho máximo de cada coluna
      jsonArray.forEach((row) => {
        headers.forEach((header) => {
          const columnLength = String(row[header]).length;
          if (!columnMaxLength[header] || columnLength > columnMaxLength[header]) {
            columnMaxLength[header] = columnLength;
          }
        });
      });

      // Escreve os cabeçalhos das colunas no arquivo XLS e formatação
      headers.forEach((header, colIndex) => {
        const cell = ws.cell(1, colIndex + 1);
        cell.string(header);
        cell.style({
          font: { bold: true },
          alignment: { horizontal: 'center' },
        });
        ws.column(colIndex + 1).setWidth(columnMaxLength[header] + 5); // Ajusta o tamanho da coluna
      });

      // Escreve o conteúdo do CSV no arquivo XLS a partir da segunda linha
      jsonArray.forEach((row, rowIndex) => {
        headers.forEach((header, colIndex) => {
          ws.cell(rowIndex + 2, colIndex + 1).string(row[header]).style(style);
        });
      });
    }

    const xlsFileName = fileName || 'converted_file.xlsx'; // Usa o nome fornecido ou um nome padrão
    const filePath = path.join(__dirname, 'src', xlsFileName);

    wb.write(filePath, (err, stats) => {
      if (err) {
        return res.status(500).json({ error: 'Erro ao converter para XLS.' });
      }

      if (convertedFilePath) {
        fs.unlinkSync(convertedFilePath); // Exclui o arquivo anterior, se existir
      }

      convertedFilePath = filePath;

      const fileUrl = `/download/${xlsFileName}`; // URL para download do arquivo

      res.json({ fileUrl });
    });
  } catch (error) {
    res.status(500).json({ error: 'Erro interno do servidor.' });
  }
});

// Rota para download do arquivo convertido
app.get('/download/:fileName', (req, res) => {
  const { fileName } = req.params;
  const filePath = path.join(__dirname, 'src', fileName);

  if (fs.existsSync(filePath)) {
    res.download(filePath);
  } else {
    res.status(404).json({ error: 'Arquivo não encontrado.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
