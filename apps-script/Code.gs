// Google Apps Script — cole isto no editor do Apps Script vinculado à planilha
// (Extensões → Apps Script), substituindo o conteúdo atual, e reimplante
// ("Implantar → Gerenciar implantações → editar → Nova versão").
//
// Cabeçalhos esperados na linha 1 da planilha (nessa ordem ou em qualquer
// ordem — a comparação ignora maiúsculas/minúsculas e espaços nas pontas):
//   Nome | Telefone | Email | Forma de Pagamento | Data | Order NSU | Compra confirmada
//
// "Compra confirmada" (e "Livro retirado", se existir) devem ser colunas
// formatadas como caixa de seleção (checkbox) — o script escreve
// true/false nelas, não texto.

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

function norm(s) {
  return String(s || "").trim().toLowerCase();
}

function getSheetAndHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return { sheet, headers };
}

function colIndex(headers, name) {
  const target = norm(name);
  const idx = headers.findIndex(h => norm(h) === target);
  return idx === -1 ? null : idx + 1; // getRange usa índice base 1
}

function handleLead(data) {
  const { sheet, headers } = getSheetAndHeaders();
  const row = headers.map(h => {
    switch (norm(h)) {
      case "nome": return data.name || "";
      case "telefone": return data.phone || "";
      case "email":
      case "e-mail": return data.email || "";
      case "forma de pagamento": return data.paymentMethod || "";
      case "data": return data.timestamp || "";
      case "order nsu": return data.orderNsu || "";
      case "compra confirmada": return false;
      case "livro retirado": return false;
      default: return "";
    }
  });
  sheet.appendRow(row);
}

function handlePaymentWebhook(data) {
  const { sheet, headers } = getSheetAndHeaders();
  const orderCol = colIndex(headers, "Order NSU");
  const confirmCol = colIndex(headers, "Compra confirmada");
  const methodCol = colIndex(headers, "Forma de Pagamento");
  const emailCol = colIndex(headers, "Email") || colIndex(headers, "E-mail");
  const nameCol = colIndex(headers, "Nome");

  if (!orderCol) return; // planilha sem a coluna "Order NSU" — não dá pra casar o pedido

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][orderCol - 1]) === String(data.order_nsu)) {
      const rowNum = i + 1;
      if (confirmCol) sheet.getRange(rowNum, confirmCol).setValue(true);
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
