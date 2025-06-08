
/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("generateBtn")!.onclick = generateEmail;
    document.getElementById("insertBtn")!.onclick = insertEmailToCompose;
    document.getElementById("regenerateBtn")!.onclick = generateEmail;
  }
});

async function generateEmail(): Promise<void> {
  try {
    showLoading(true);
    hideResult();
    
    const subject = (document.getElementById("emailSubject") as HTMLInputElement).value.trim();
    const recipient = (document.getElementById("emailRecipient") as HTMLInputElement).value.trim();
    const context = (document.getElementById("emailContext") as HTMLTextAreaElement).value.trim();
    const tone = (document.getElementById("emailTone") as HTMLSelectElement).value;
    
    if (!subject || !recipient || !context) {
      alert("件名、宛先、メール内容をすべて入力してください。");
      return;
    }
    
    const emailDraft = generateEmailTemplate(subject, recipient, context, tone);
    showResult(emailDraft);
    
  } catch (error) {
    console.error('Error generating email:', error);
    alert(`エラーが発生しました: ${error.message}`);
  } finally {
    showLoading(false);
  }
}

function generateEmailTemplate(subject: string, recipient: string, context: string, tone: string): string {
  const templates = {
    formal: `${recipient}

いつもお世話になっております。

${subject}の件について、ご連絡申し上げます。

${context}

ご不明な点等ございましたら、お気軽にお申し付けください。
今後ともよろしくお願い申し上げます。

──────────────────
［お名前］
［会社名］
［連絡先］`,

    business: `${recipient}

お疲れさまです。

${subject}について、ご連絡いたします。

${context}

ご確認のほど、よろしくお願いいたします。

──────────────────
［お名前］`,

    casual: `${recipient}

${context}

よろしくお願いします！

──────────────────
［お名前］`
  };

  return templates[tone] || templates.business;
}

function insertEmailToCompose(): void {
  const generatedEmail = document.getElementById("generated-email")!.textContent;
  
  if (!generatedEmail) {
    alert("挿入するメール内容がありません。");
    return;
  }
  
  Office.context.mailbox.item.body.setAsync(
    generatedEmail,
    { coercionType: Office.CoercionType.Text },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        alert("メールが正常に挿入されました！");
      } else {
        alert("メールの挿入に失敗しました。");
      }
    }
  );
}

function showLoading(show: boolean): void {
  const loading = document.getElementById("loading")!;
  loading.style.display = show ? "block" : "none";
}

function showResult(emailContent: string): void {
  const resultSection = document.getElementById("result-section")!;
  const generatedEmail = document.getElementById("generated-email")!;
  
  generatedEmail.textContent = emailContent;
  resultSection.style.display = "block";
}

function hideResult(): void {
  const resultSection = document.getElementById("result-section")!;
  resultSection.style.display = "none";
}