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

const EBOOK_FILE_ID = "1kcBt14Hy-3Iiui9_qWRT_hzRjcGMBZr6"; // E-book Livre da Pornografia.pdf
const SEND_FROM_ALIAS = "construindoaproximageracao@gmail.com"; // precisa estar verificado em "Enviar e-mail como" na conta dona do script

// TESTE MANUAL: selecione "testEnvioManual" no menu suspenso ao lado do
// botão "Depurar/Executar" no topo do editor, e clique em "Executar".
// O erro (se houver) aparece direto aqui no editor, sem precisar do
// Cloud Logging. Troque o e-mail abaixo pelo seu antes de rodar.
function testEnvioManual() {
  sendEbookEmail("guibaand@gmail.com", "Teste Manual");
  Logger.log("testEnvioManual: terminou sem lançar erro.");
}

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

function buildEbookEmailHtml(name) {
  const firstName = (name || "").trim().split(" ")[0] || "";
  const greeting = firstName ? `Parabéns, ${firstName}!` : "Parabéns!";

  return `
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Seu exemplar está garantido</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Anton&family=Nunito+Sans:wght@400;600;700;800&display=swap">
<style>
  @media (max-width: 640px) {
    .bcpg-email-title { font-size: 24px !important; }
  }
</style>
</head>
<body style="margin:0;padding:0;background-color:#F4EDDB;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#F4EDDB;padding:32px 16px;">
  <tr>
    <td align="center">
      <table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background-color:#FFFDF6;border-radius:20px;overflow:hidden;">

        <tr>
          <td style="background-color:#111110;padding:24px 32px;">
            <table role="presentation" cellpadding="0" cellspacing="0">
              <tr>
                <td style="width:12px;height:28px;background-color:#E43242;border-radius:100px;font-size:0;line-height:0;">&nbsp;</td>
                <td style="padding-left:12px;font-family:'Anton',Impact,'Arial Black',sans-serif;font-size:16px;font-weight:700;letter-spacing:.04em;line-height:1.2;color:#F4EDDB;">
                  CONSTRUINDO<br>
                  <span style="color:#E43242;">A PRÓXIMA GERAÇÃO</span>
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <tr>
          <td style="padding:40px 32px 8px;">
            <div style="font-size:12px;font-weight:800;letter-spacing:.16em;text-transform:uppercase;color:#E43242;font-family:'Nunito Sans',Arial,Helvetica,sans-serif;">
              Compra confirmada
            </div>
            <div class="bcpg-email-title" style="font-family:'Anton',Impact,'Arial Black',sans-serif;font-size:32px;line-height:1.1;font-weight:700;text-transform:uppercase;color:#111110;margin-top:12px;">
              ${greeting}<br>Seu exemplar está garantido.
            </div>
          </td>
        </tr>

        <tr>
          <td style="padding:16px 32px 0;font-family:'Nunito Sans',Arial,Helvetica,sans-serif;font-size:16px;line-height:1.6;color:#3A3730;">
            Recebemos a confirmação do seu pagamento. Seu exemplar de <strong>Construindo a Próxima Geração</strong>, de Pedro Lacava, está reservado — é só retirar presencialmente no dia <strong>26 de setembro</strong>, no <strong>A13 — CCVIDEIRA CANDELÁRIA</strong>.
          </td>
        </tr>

        <tr>
          <td style="padding:24px 32px 0;">
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#F4EDDB;border-radius:16px;">
              <tr>
                <td style="padding:24px;">
                  <div style="font-family:'Nunito Sans',Arial,Helvetica,sans-serif;font-size:12px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#E43242;">
                    Bônus da pré-venda
                  </div>
                  <div style="font-family:'Anton',Impact,'Arial Black',sans-serif;font-size:20px;font-weight:700;color:#111110;margin-top:8px;text-transform:uppercase;">
                    E-book: Livre-se da Pornografia
                  </div>
                  <div style="font-family:'Nunito Sans',Arial,Helvetica,sans-serif;font-size:15px;line-height:1.6;color:#3A3730;margin-top:8px;">
                    Segue em anexo a este e-mail, em PDF. É seu, de presente, como agradecimento por fazer parte da pré-venda.
                  </div>
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <tr>
          <td style="padding:32px 32px 8px;font-family:'Nunito Sans',Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#6E6A5E;">
            Qualquer dúvida sobre a retirada, é só chamar no WhatsApp <strong style="color:#111110;">(84) 9683-0709</strong>.
          </td>
        </tr>

        <tr>
          <td style="padding:8px 32px 32px;font-family:'Nunito Sans',Arial,Helvetica,sans-serif;font-size:14px;line-height:1.6;color:#3A3730;">
            Deus abençoe!<br>
            <strong>Pedro Lacava</strong>
          </td>
        </tr>

        <tr>
          <td style="background-color:#111110;padding:20px 32px;font-family:'Nunito Sans',Arial,Helvetica,sans-serif;font-size:12px;color:rgba(244,237,219,.6);">
            © 2026 Construindo a Próxima Geração. Todos os direitos reservados.
          </td>
        </tr>

      </table>
    </td>
  </tr>
</table>
</body>
</html>`;
}

function sendEbookEmail(email, name) {
  if (!EBOOK_FILE_ID || !email) {
    Logger.log("sendEbookEmail: pulado (EBOOK_FILE_ID ou email ausente). email=" + email);
    return;
  }

  Logger.log("sendEbookEmail: iniciando envio para " + email);

  const file = DriveApp.getFileById(EBOOK_FILE_ID);
  const options = {
    htmlBody: buildEbookEmailHtml(name),
    attachments: [file.getAs(MimeType.PDF)],
    name: "Construindo a Próxima Geração"
  };
  if (SEND_FROM_ALIAS) options.from = SEND_FROM_ALIAS;

  GmailApp.sendEmail(email, "Seu exemplar está garantido + e-book bônus", "", options);
  Logger.log("sendEbookEmail: GmailApp.sendEmail retornou sem erro para " + email + " (from=" + (SEND_FROM_ALIAS || "conta padrão") + ")");
  Logger.log("Cota restante de e-mail hoje: " + MailApp.getRemainingDailyQuota());
}
