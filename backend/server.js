// backend/server.js 

const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
// const nodemailer = require('nodemailer');
// require('dotenv').config();

console.log('MJ_USERNAME:', process.env.MJ_USERNAME ? 'OK' : 'FALTANDO!');
console.log('MJ_PASSWORD:', process.env.MJ_PASSWORD ? 'OK' : 'FALTANDO!');

if (!process.env.MJ_USERNAME || !process.env.MJ_PASSWORD) {
  console.error('ERRO: Variáveis do Mailjet não encontradas no .env');
  process.exit(1);
}

const Mailjet = require('node-mailjet');
const mailjet = new Mailjet({
  apiKey: process.env.MJ_USERNAME,
  apiSecret: process.env.MJ_PASSWORD
});

const path = require('path');
const fs = require('fs');

const app = express();

const cors = require('cors');
app.use(cors());

// === PASTA PARA ARQUIVOS TEMPORÁRIOS ===
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);

// === MÚLTIPLOS ARQUIVOS ===
const upload = multer({ dest: uploadDir, limits: { fileSize: 15 * 1024 * 1024 } });

app.use(express.json());




// === ROTA PRINCIPAL ===
app.post('/enviar', upload.array('arquivos'), async (req, res) => {
  try {
    // === DADOS DO USUÁRIO ===
    const { nome, email, cpf, banco, agencia, conta, pix, project, dataAtual } = req.body;
    const arquivos = req.files || [];
    const nomeArquivo = arquivos.length > 0 ? arquivos[0].originalname : '';

    console.log('Dados do usuário:', { nome, email, cpf, arquivos: arquivos.map(f => f.originalname) });

    // === catch DESPESAS ===
    let despesas = [];
    const despesasJson = req.body.despesas;
    if (despesasJson) {
      try {
        despesas = JSON.parse(despesasJson);
      } catch (e) {
        return res.status(400).json({ success: false, message: 'Erro ao ler despesas' });
      }
    }



    // === ASSOCIAR ARQUIVOS NA ORDEM ===
    arquivos.forEach((file, i) => {
      if (despesas[i]) {
        despesas[i].arquivoNome = file.originalname;
        despesas[i].arquivoPath = file.path;
      }
    });

    console.log('Despesas finais:', despesas);



    // === GERAR PLANILHA ===
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Expense Report', {
      pageSetup: { paperSize: 9, orientation: 'landscape' },
      views: [{ showGridLines: false }]
    });

    // === FUNÇÃO DE BORDA ===
    const applyBorder = (cell, options = {}) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
        ...options
      };
    };


    const applyBorder2 = (cell, options = {}) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        ...options
      };
    };

    // === LOGO ===
    sheet.mergeCells('A1:C4');
    const logoCell = sheet.getCell('A1');
    logoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

    // === bordas ===
    ['A4', 'B4', 'C4'].forEach(c => {
      const cell = sheet.getCell(c);
      cell.border = cell.border || {};
      cell.border.bottom = { style: 'thin' };
    });

    ['C1', 'C2', 'C3', 'C4'].forEach(c => {
      const cell = sheet.getCell(c);
      cell.border = cell.border || {};
      cell.border.right = { style: 'thin' };
    });



    const logoPath = path.join(__dirname, 'logo.jpg');
    if (fs.existsSync(logoPath)) {
      const logoId = workbook.addImage({ filename: logoPath, extension: 'png' });
      sheet.addImage(logoId, {
        tl: { col: 0, row: 0 },
        ext: { width: 258, height: 78 },
        editAs: 'oneCell'
      });
    }
    sheet.getColumn('A').width = 23;



    // === FUNDO D1:N4 ===
    for (let r = 1; r <= 4; r++) {
      for (let c = 4; c <= 14; c++) {
        const cell = sheet.getCell(r, c);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2DCDB' } };
      }
    }

    // === TÍTULO ===
    sheet.mergeCells('D1:N1');
    const title = sheet.getCell('D1');
    title.value = 'EXPENSE REPORT FOR REIMBURSEMENT';
    title.font = { size: 16, bold: true };
    title.alignment = { horizontal: 'center', vertical: 'middle' };
    applyBorder(title);

    for (let c = 9; c <= 14; c++) {
      const cell = sheet.getCell(1, c);
      cell.border = cell.border || {};
      cell.border.bottom = { style: 'thin' };
    }

    // === DADOS DO FUNCIONÁRIO ===
    const dados = [
      ['EMPLOYEE NAME:', nome],
      ['CPF:', cpf],
      ['EMAIL:', email]
    ];
    dados.forEach(([label, valor], i) => {
      const row = 2 + i;
      sheet.getCell(`D${row}`).value = label;
      sheet.getCell(`D${row}`).font = { bold: true };
      sheet.getCell(`E${row}`).value = valor || '';
      for (let c = 4; c <= 8; c++) {
        applyBorder(sheet.getCell(row, c));
      }
    });

    sheet.getCell('G2').value = 'Project Number / Cost Center:'; sheet.getCell('G2').font = { bold: true };
    sheet.getCell('H2').value = project || '';
    sheet.getCell('G3').value = 'Date:'; sheet.getCell('G3').font = { bold: true };
    sheet.getCell('H3').value = dataAtual || '';

    for (let c = 4; c <= 14; c++) {
      const cell = sheet.getCell(4, c);
      cell.border = cell.border || {};
      cell.border.bottom = { style: 'thin' };
    }
    sheet.getCell('D4').border.right = { style: 'thin' };

    for (let r = 1; r <= 4; r++) {
      const cell = sheet.getCell(r, 14);
      cell.border = cell.border || {};
      cell.border.right = { style: 'thin' };
    }

    // === TAMANHO DAS COLUNAS ===
    sheet.getColumn('D').width = 17;
    sheet.getColumn('E').width = 30;
    sheet.getColumn('G').width = 28;
    sheet.getColumn('H').width = 12;
    sheet.getColumn('K').width = 13;
    sheet.getColumn('L').width = 13;
    sheet.getColumn('N').width = 11;
    sheet.getColumn('M').width = 23;
    sheet.getColumn('B').width = 10;

    // === CABEÇALHO DA TABELA ===
    const headers = [
      'DOC', 'DATE', 'ACTIVITY', 'LOCAL', 'DESCRIPTION',
      'TRANSPORT TOLL', 'HOTEL', 'MEALS', 'GASOLINE', 'OTHER'
    ];
    headers.forEach((h, i) => {
      const col = String.fromCharCode(65 + i);
      sheet.mergeCells(`${col}6:${col}7`);
      const cell = sheet.getCell(`${col}6`);
      cell.value = h;
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2DCDB' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      applyBorder(cell);
    });

    // ONLY INTERNATIONAL
    sheet.mergeCells('K6:L6');
    const onlyInt = sheet.getCell('K6');
    onlyInt.value = 'ONLY INTERNATIONAL';
    onlyInt.font = { bold: true };
    onlyInt.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2DCDB' } };
    onlyInt.alignment = { horizontal: 'center', vertical: 'middle' };
    applyBorder(onlyInt);

    // FOREIGN CURRENCY + EXCHANGE RATE
    sheet.getCell('K7').value = 'FOREIGN\nCURRENCY';
    sheet.getCell('L7').value = 'EXCHANGE\nRATE';
    ['K7', 'L7'].forEach(c => {
      const cell = sheet.getCell(c);
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2DCDB' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      applyBorder(cell);
      cell.border.right = { style: 'thin' };
    });


    // REFERÊNCIA
    sheet.mergeCells('M6:M7');
    const refHeader = sheet.getCell('M6');
    refHeader.value = 'REFERÊNCIA';
    refHeader.font = { bold: true };
    refHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2DCDB' } };
    refHeader.alignment = { horizontal: 'center', vertical: 'middle' };
    applyBorder(refHeader);


    // TOTAL (BRL)
    sheet.mergeCells('N6:N7');
    const totalHeader = sheet.getCell('N6');
    totalHeader.value = 'TOTAL (BRL)';
    totalHeader.font = { bold: true };
    totalHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2DCDB' } };
    totalHeader.alignment = { horizontal: 'center', vertical: 'middle' };
    applyBorder(totalHeader);

    // === LINHAS DAS DESPESAS ===
    let rowIndex = 8;
    let subtotal = 0;
    const colTotals = Array(14).fill(0); // até N

    despesas.forEach(d => {
      const row = sheet.getRow(rowIndex);
      const values = [
        d.doc || '',
        d.data || '',
        d.atividade || '',
        d.local || '',
        d.descricao || '',
        d.pedagio ? parseFloat(d.pedagio) : 0,
        d.hotel ? parseFloat(d.hotel) : 0,
        d.refeicoes ? parseFloat(d.refeicoes) : 0,
        d.gasolina ? parseFloat(d.gasolina) : 0,
        d.other ? parseFloat(d.other) : 0,
        d.moeda || '',
        d.taxa ? parseFloat(d.taxa) : '',
        d.arquivoNome || '',
        0 // TOTAL (BRL)
      ];

      const totalLinha = values.slice(5, 10).reduce((a, b) => a + (b || 0), 0);
      values[13] = totalLinha; // coluna N
      subtotal += totalLinha;

      row.values = values;
      row.eachCell({ includeEmpty: true }, (cell, col) => {
        applyBorder(cell);
        if (col >= 6 && col <= 10 || col === 14) {
          cell.numFmt = '"R$ "#,##0.00';
          colTotals[col - 1] += cell.value || 0;
        }
      });
      rowIndex++;
    });

    // === SUBTOTAL (AMARELO) ===
    const subtotalRow = sheet.getRow(rowIndex++);
    sheet.mergeCells(`A${subtotalRow.number}:E${subtotalRow.number}`);
    subtotalRow.getCell(1).value = 'Subtotal';
    subtotalRow.getCell(1).font = { bold: true };

    for (let c = 1; c <= 5; c++) applyBorder(subtotalRow.getCell(c));

    for (let c = 6; c <= 10; c++) {
      subtotalRow.getCell(c).value = colTotals[c - 1];
      subtotalRow.getCell(c).numFmt = '"R$ "#,##0.00';
      subtotalRow.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
      applyBorder(subtotalRow.getCell(c));
    }

    sheet.mergeCells(`K${subtotalRow.number}:M${subtotalRow.number}`);
    const mergedKLM = subtotalRow.getCell(11);
    mergedKLM.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF' } };
    mergedKLM.alignment = { horizontal: 'center', vertical: 'middle' };
    applyBorder(mergedKLM);
    

    // TOTAL (BRL)
    subtotalRow.getCell(14).value = subtotal;
    subtotalRow.getCell(14).numFmt = '"R$ "#,##0.00';
    subtotalRow.getCell(14).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    applyBorder(subtotalRow.getCell(14));

    // === TOTAIS FINAIS ===
    const totalStart = rowIndex + 1;
    const totais = [
      { label: 'SUBTOTAL:', value: subtotal, color: 'FFFFFF00' },
      { label: 'ADVANCES:', value: 0, color: 'FFFFA500' },
      { label: 'TOTAL REIMBURSEMENT:', value: subtotal, color: 'FF90EE90' }
    ];

    totais.forEach((t, i) => {
      const r = totalStart + i;
      sheet.getCell(`M${r}`).value = t.label;
      sheet.getCell(`M${r}`).font = { bold: true };
      sheet.getCell(`M${r}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: t.color } };
      applyBorder(sheet.getCell(`M${r}`));

      sheet.getCell(`N${r}`).value = t.value;
      sheet.getCell(`N${r}`).numFmt = '"R$ "#,##0.00';
      sheet.getCell(`N${r}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: t.color } };
      applyBorder(sheet.getCell(`N${r}`));
    });

    // === ASSINATURAS ===
    const empSigStart = rowIndex + 3;

    const createSignature = (startRow, startCol, endCol, title, bgColor) => {
      for (let r = startRow; r < startRow + 3; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const cell = sheet.getCell(r, c);
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
          cell.border = {};
          if (r === startRow) cell.border.top = { style: 'thin' };
          if (r === startRow + 2) cell.border.bottom = { style: 'thin' };
          if (c === startCol) cell.border.left = { style: 'thin' };
          if (c === endCol) cell.border.right = { style: 'thin' };
        }
      }
      sheet.getCell(startRow, startCol).value = title;
      sheet.getCell(startRow, startCol).font = { bold: true };
      const dateCell = sheet.getCell(startRow + 1, endCol);
      dateCell.value = 'DATE';
      dateCell.font = { bold: true };
      applyBorder(dateCell);
      applyBorder(sheet.getCell(startRow + 2, endCol));
    };

    createSignature(empSigStart, 1, 4, 'EMPLOYEE SIGNATURE', 'FFF2F2F2');
    createSignature(empSigStart + 3, 1, 4, 'DIRECT MANAGER APPROVAL NAME AND SIGNATURE', 'FFDCE6F1');
    createSignature(empSigStart + 3, 5, 8, 'BUSINESS UNIT MANAGER APPROVAL NAME AND SIGNATURE', 'FFDCE6F1');

    // === DADOS BANCÁRIOS ===
    const bankStartRow = empSigStart;
    const bankStartCol = 5;
    const bankEndCol = 8;

    for (let r = bankStartRow; r < bankStartRow + 3; r++) {
      for (let c = bankStartCol; c <= bankEndCol; c++) {
        const cell = sheet.getCell(r, c);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
        cell.border = {};
        if (r === bankStartRow) cell.border.top = { style: 'thin' };
        if (r === bankStartRow + 2) cell.border.bottom = { style: 'thin' };
        if (c === bankStartCol) cell.border.left = { style: 'thin' };
        if (c === bankEndCol) cell.border.right = { style: 'thin' };
      }
    }

    sheet.getCell(bankStartRow, bankStartCol).value = `Bank: ${banco || ''}`;
    sheet.getCell(bankStartRow + 1, bankStartCol).value = `Agency: ${agencia || ''}`;
    sheet.getCell(bankStartRow + 2, bankStartCol).value = `Account: ${conta || ''}`;
    sheet.getCell(bankStartRow + 1, bankStartCol + 2).value = `PIX: ${pix || ''}`;

    // === SALVAR ===
    const excelPath = path.join(__dirname, 'Expense_Report.xlsx');
    await workbook.xlsx.writeFile(excelPath);



    // === ANEXOS ===
    // const attachments = [{ filename: 'despesas.xlsx', path: excelPath }];
    // despesas.forEach(d => {
    //   if (d.arquivoPath) attachments.push({ filename: d.arquivoNome, path: d.arquivoPath });
    // });

    // === EMAIL ===
    // const transporter = nodemailer.createTransport({
    //   host: 'in-v3.mailjet.com',
    //   port: 465,
    //   secure: true,
    //   auth: {
    //     user: process.env.MJ_USERNAME, // API Key
    //     pass: process.env.MJ_PASSWORD  // Secret Key
    //   }
    // });

    // // Corpo do email em HTML
    // const htmlBody = `
    //   <h2>Nova solicitação de Reembolso de Despesas</h2>
    //   <p><strong>Detalhes do solicitante:</strong></p>
    //   <ul>
    //     <li>Nome: ${nome}</li>
    //     <li>Email: ${email}</li>
    //     <li>Data do formulário: ${dataAtual}</li>
    //   </ul>
    //   <p><strong>Resumo das Despesas:</strong></p>
    //   <p>Total de Despesas: <strong>${despesas.length}</strong></p>
    //   <p>Planilha com detalhes e comprovantes em anexo.</p>
    // `;

    // await transporter.sendMail({
    //   from: '"Kraftone Despesas" <jonathan.lemos@kraftonegroup.com>',
    //   // from: process.env.EMAIL_USER,
    //   to: 'jonathan.lemos@kraftonegroup.com',
    //   cc: email,
    //   subject: `Solicitação de Reembolso de Despesas`,
    //   html: htmlBody,
    //   attachments
    // });



        // === EMAIL COM API REST DO MAILJET ===
    const htmlBody = `
      <h2>Nova solicitação de Reembolso de Despesas</h2>
      <p><strong>Detalhes do solicitante:</strong></p>
      <ul>
        <li>Nome: ${nome}</li>
        <li>Email: ${email}</li>
        <li>Data do formulário: ${dataAtual}</li>
      </ul>
      <p><strong>Resumo das Despesas:</strong></p>
      <p>Total de Despesas: <strong>${despesas.length}</strong></p>
      <p>Planilha com detalhes e comprovantes em anexo.</p>
    `;

    // LER PLANILHA EM BASE64
    const planilhaBuffer = fs.readFileSync(excelPath);
    const planilhaBase64 = planilhaBuffer.toString('base64');

    // PREPARAR ANEXOS
    const attachments = [
      {
        ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        Filename: `despesas_${nome}.xlsx`,
        Base64Content: planilhaBase64
      }
    ];

    // ADICIONAR ARQUIVOS DO USUÁRIO
    arquivos.forEach(file => {
      const fileBuffer = fs.readFileSync(file.path);
      attachments.push({
        ContentType: file.mimetype,
        Filename: file.originalname,
        Base64Content: fileBuffer.toString('base64')
      });
    });

    // ENVIAR COM API REST (V6.0.9)
    try {
      const request = await mailjet.post('send', { version: 'v3.1' }).request({
        Messages: [
          {
            From: {
              Email: 'hotline.kraftonegroup@gmail.com',
              Name: 'Kraftone Despesas'
            },
            To: [{ Email: email, Name: nome }],
            Cc: [{ Email: 'jonathan.lemos@kraftonegroup.com' }],
            Subject: `Solicitação de Reembolso - ${nome}`,
            HTMLPart: htmlBody,
            Attachments: attachments
          }
        ]
      });

      console.log('Email enviado via API REST!');
    } catch (mailError) {
      console.error('ERRO AO ENVIAR EMAIL:', mailError);
      throw mailError;
    }



    // === LIMPAR ARQUIVOS TEMPORÁRIOS ===
    fs.unlinkSync(excelPath);
    arquivos.forEach(f => fs.unlinkSync(f.path));

    // === RESPOSTA ===
    res.set('Content-Type', 'application/json');
    return res.status(200).send(JSON.stringify({
      success: true,
      message: 'Enviado com sucesso!'
    }));

  } catch (error) {
    console.error('ERRO:', error);
    const erro = JSON.stringify({ success: false, message: 'Erro no servidor' });
    res.writeHead(500, {
      'Content-Type': 'application/json',
      'Content-Length': Buffer.byteLength(erro)
    });
    res.end(erro);
  }

});

// === INICIAR SERVIDOR ===
app.listen(3001, () => console.log('Backend rodando na porta 3001'));