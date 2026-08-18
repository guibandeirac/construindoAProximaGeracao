// Google Apps Script — cole isto no editor do Apps Script vinculado à planilha
// (Extensões → Apps Script), substituindo o conteúdo atual, e reimplante
// ("Implantar → Gerenciar implantações → editar → Nova versão").
//
// Cabeçalhos esperados na linha 1 da planilha (nessa ordem ou em qualquer
// ordem, desde que os nomes sejam exatamente estes):
//   Nome | Telefone | Email | Forma de pagamento | Data | Order NSU | Compra confirmada

const EBOOK_FILE_ID = ""; // TODO: cole aqui o ID do arquivo PDF do e-book no Google Drive

function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  // A InfinitePay manda esses campos quando o pagamento é aprovado.
  // Nosso próprio site nunca manda "invoice_slug" nem "transaction_nsu".
  if (body.invoice_slug || body.transaction_nsu) {
    handlePaymentWebhook(body);
  } else {
    handleLead(body);
  }

  return ContentService.createTextOutput("ok");
}

function getSheetAndHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return { sheet, headers };
}

function colIndex(headers, name) {
  const idx = headers.indexOf(name);
  return idx === -1 ? null : idx + 1; // getRange usa índice base 1
}

function handleLead(data) {
  const { sheet, headers } = getSheetAndHeaders();
  const row = headers.map(h => {
    switch (h) {
      case "Nome": return data.name || "";
      case "Telefone": return data.phone || "";
      case "Email":
      case "E-mail": return data.email || "";
      case "Forma de pagamento": return data.paymentMethod || "";
      case "Data": return data.timestamp || "";
      case "Order NSU": return data.orderNsu || "";
      case "Compra confirmada": return "Não";
      default: return "";
    }
  });
  sheet.appendRow(row);
}

function handlePaymentWebhook(data) {
  const { sheet, headers } = getSheetAndHeaders();
  const orderCol = colIndex(headers, "Order NSU");
  const confirmCol = colIndex(headers, "Compra confirmada");
  const methodCol = colIndex(headers, "Forma de pagamento");
  const emailCol = colIndex(headers, "Email") || colIndex(headers, "E-mail");
  const nameCol = colIndex(headers, "Nome");

  if (!orderCol) return; // planilha sem a coluna "Order NSU" — não dá pra casar o pedido

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][orderCol - 1]) === String(data.order_nsu)) {
      const rowNum = i + 1;
      if (confirmCol) sheet.getRange(rowNum, confirmCol).setValue("Sim");
      if (methodCol) {
        sheet.getRange(rowNum, methodCol).setValue(
          data.capture_method === "pix" ? "Pix" : "Cartão de crédito"
        );
      }
      if (emailCol && nameCol) {
        sendEbookEmail(values[i][emailCol - 1], values[i][nameCol - 1]);
      }
      break;
    }
  }
}

function sendEbookEmail(email, name) {
  if (!EBOOK_FILE_ID || !email) return; // PDF ainda não configurado

  const file = DriveApp.getFileById(EBOOK_FILE_ID);
  const htmlBody = `
    <p>Olá${name ? ", " + name : ""}!</p>
    <p>Sua compra do livro <strong>Construindo a Próxima Geração</strong> foi confirmada. Como agradecimento, segue seu e-book bônus <strong>"Livre-se da Pornografia"</strong> em anexo.</p>
    <p>Qualquer dúvida, é só responder este e-mail ou chamar no WhatsApp (84) 9683-0709.</p>
    <p>Deus abençoe!<br>Pedro Lacava</p>
  `;

  GmailApp.sendEmail(email, "Seu e-book bônus: Livre-se da Pornografia", "", {
    htmlBody: htmlBody,
    attachments: [file.getAs(MimeType.PDF)],
    name: "Construindo a Próxima Geração"
  });
}
