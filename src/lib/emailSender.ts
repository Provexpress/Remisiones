export interface SendEmailPayload {
  toEmail: string;
  toName: string;
  subject: string;
  htmlBody: string;
  excelAttachment?: {
    filename: string;
    base64: string;
  };
}

export interface SendResult {
  success: boolean;
  toEmail: string;
  toName: string;
  error?: string;
}

import {
  LOGO_PROVEXPRESS_BASE64,
  AVATAR_MAN_BASE64,
  AVATAR_WOMAN_BASE64,
} from './commercialEmailAssets';

/**
 * Envía un correo electrónico a través de Microsoft Graph API (/me/sendMail).
 * Utiliza el token de autenticación de Microsoft 365.
 */
export async function sendMailViaGraph(
  accessToken: string,
  payload: SendEmailPayload,
): Promise<SendResult> {
  const endpoint = 'https://graph.microsoft.com/v1.0/me/sendMail';

  const attachments: any[] = [
    {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'logo_provexpress.png',
      contentType: 'image/png',
      contentBytes: LOGO_PROVEXPRESS_BASE64,
      isInline: true,
      contentId: 'logo_provexpress',
    },
  ];

  if (payload.htmlBody.includes('cid:avatar_woman')) {
    attachments.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'avatar_woman.png',
      contentType: 'image/png',
      contentBytes: AVATAR_WOMAN_BASE64,
      isInline: true,
      contentId: 'avatar_woman',
    });
  }

  if (payload.htmlBody.includes('cid:avatar_man')) {
    attachments.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'avatar_man.png',
      contentType: 'image/png',
      contentBytes: AVATAR_MAN_BASE64,
      isInline: true,
      contentId: 'avatar_man',
    });
  }

  if (payload.excelAttachment) {
    attachments.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: payload.excelAttachment.filename,
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      contentBytes: payload.excelAttachment.base64,
      isInline: false,
    });
  }

  const body = {
    message: {
      subject: payload.subject,
      body: {
        contentType: 'HTML',
        content: payload.htmlBody,
      },
      toRecipients: [
        {
          emailAddress: {
            address: payload.toEmail,
            name: payload.toName,
          },
        },
      ],
      attachments,
    },
    saveToSentItems: true,
  };

  try {
    const response = await fetch(endpoint, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const errData = await response.json().catch(() => ({}));
      const msg = errData?.error?.message || `Error ${response.status}: ${response.statusText}`;
      return {
        success: false,
        toEmail: payload.toEmail,
        toName: payload.toName,
        error: msg,
      };
    }

    return {
      success: true,
      toEmail: payload.toEmail,
      toName: payload.toName,
    };
  } catch (err: any) {
    return {
      success: false,
      toEmail: payload.toEmail,
      toName: payload.toName,
      error: err?.message || 'Error de conexión al enviar el correo.',
    };
  }
}

/**
 * Envío masivo con control de tasa (rate limiting) para Microsoft 365 Exchange Online.
 * Incluye callback de progreso para la interfaz de usuario.
 */
export async function sendBatchEmails(
  accessToken: string,
  items: SendEmailPayload[],
  onProgress?: (current: number, total: number, lastResult: SendResult) => void,
  delayMs = 600,
): Promise<SendResult[]> {
  const results: SendResult[] = [];

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const res = await sendMailViaGraph(accessToken, item);
    results.push(res);

    if (onProgress) {
      onProgress(i + 1, items.length, res);
    }

    // Pausa preventiva entre correos para respetar las cuotas de Microsoft Graph
    if (i < items.length - 1 && delayMs > 0) {
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    }
  }

  return results;
}
