import type { VercelRequest, VercelResponse } from '@vercel/node';
import { Resend } from 'resend';

// ==================== 設定 ====================
const LARK_APP_ID     = process.env.LARK_APP_ID     || 'cli_a95e55ea22a19e18';
const LARK_APP_SECRET = process.env.LARK_APP_SECRET || '77R4UsD3V9yoIzJKaFwcdhtpQl5gsOUa';
const LARK_TABLE_ID   = process.env.LARK_TABLE_ID   || 'tblsi0HfqNtxj46W';
const RESEND_API_KEY  = process.env.RESEND_API_KEY  || 're_749biiA7_7ZChV5CUbRzexeHM5S6aNc1r';
const EMAIL_TO        = 'jump@pocketpro.tw';
const EMAIL_FROM      = 'PocketPro <onboarding@get-pocketpro.com>';

// Google Drive
const DRIVE_FOLDER_ID = '1DRCiiofdzzGXZeEYSpfUmdbPva84DfRA';
const SERVICE_ACCOUNT_EMAIL = 'pocketpro-sheet@project-d61f7947-3913-4800-b08.iam.gserviceaccount.com';
const SERVICE_ACCOUNT_KEY = `-----BEGIN PRIVATE KEY-----
MIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQCv44ONg1vbtDZy
Z0/nfKztYIpGxIgfvzQCkWXErg0vYXhMOYcK7sLFigp53J5An5dClGTASxXWMZlw
96GoyQib0RFHTJBX1nmUZJfinayf82+UDSItg3YkvgXynPX/9fOr51of+AaG5jWH
0EO3Y7/lnEr/RIcBih4UW65Sj7sBwIWRN3OMLIMguEnKTNLU/5pI4s46ezVXjIgS
q0m1x/JxgD8WsIQ+Ak1RKoIq4b9xKOTtmX4JRKyckz+K6Hjtn00uHHxWysGPo8T1
dQHiEXpIjS34R9E77Bh9YluTgAKBOARaTmQeLGpiPC4s/TJng5TGmlOmAYdrAMS+
gHHQZEV1AgMBAAECggEAGTEKVfni7bA9fhw67wpF0Efb9i/O2VEu11FQ1J8jJ06c
BrrUkyXIQre3MWX+Sn4xEWmkloAKlB+NfQcSodSNRZfnlCEsqVAAINdZg60WnOAm
cnuBEii6gp+uxWVivHLTICNmHp8M/EQ7lYSoNjt0sCO3ACGl/nv0O/E3of6RB7p4
+LkEG2xCMpYJR73eAAzFOU0pFeBYGEElDgWzoZO4JLI/DzfrFEsI1jI6mRkBAmEo
tbZYPGOqeKVVM7JYxFNSqIB3Z370pp8HK2Ctg+o/xtD3KuaEcR8Q3Se+GWeTZU13
i5CPq8mPzELYd1IPISz+0FKvCEI/mC0ExmNRGMUL4QKBgQDZ1nd6cYa6jfVWPc31
z5nx85R/uS8qI2DDLRvKBfsK1hy4k2MJ1e+kGbENfPri4Al31+ogT7cqMmxgNpHH
zoDAtsipy9Gz6BNk//qeUiBgWkNGLms/hzKx98QmLQlIln344MdYDMs77zWfW9o9
tFcq2ZTQeGUwSDtBDqIWwRZUhQKBgQDOs7o1nIM9GjGkscKocMqoC3585VLO7WP1
B4W3ltX7DhpYLTqfzlSonELD1LV0CCbv/Go3X07y0AzGrHUBzf4iirAhlOd1Ruqd
LRsgYsUorxASsWR50xDBySTBCPrlzYVX7y5nHCjqG4P+rGmYRzkDqAG5suodKy6n
at1mzG04MQKBgD6sBE3W8aMkimwYdfP9mVXR9WxVs+sUqJcemDskQ1iXx0WXKcw/
n6V/ur+dsHSrbi3rkbFgHdtnDGUV7hUlJUfMjqjDOf7fiwzo1IrOKABwl6BOZI6v
b/dhyC4PkPcwTOfYi6GadLI2nR/PBlfwVY+/b6AWs04TyfBqrFmNjcYdAoGAAxAA
o0i1XRNlRuZnVu2M4x6AekM/jddQktHQtl6ivvx/gWzyIGoDMRhXmOUu5xAz23xm
6nkcB1bzyYHGngc6S7K4V1cIcuFhGoEPlNRBzY+CcnR0Y6Wv6t8bD00dwofgAOSH
UHnHVWig9QYC7oGno5k6pVC0TUhVgZ+AtkQzHhECgYA+C91PBv6MxJ6A/TSUjbWh
EECnzio31464VQJn2zLNSeQp8kdfkKQ+jaBBF8c6to8HYTpEuBB6Zib4xp49ZdhC
9s4hIrq8oO51SQmYxRbWgcUrnEI6b+dI/QQjHxHVsOkC5iKsit25iFqNP7mA0Bfl
EaaIeswNF+nhhoBbOpLeeQ==
-----END PRIVATE KEY-----`;

export const dynamic = 'force-dynamic';

function setCors(res: VercelResponse) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// ==================== Google Drive ====================
async function getDriveAccessToken(): Promise<string> {
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
  const payload = Buffer.from(JSON.stringify({
    iss: SERVICE_ACCOUNT_EMAIL,
    scope: 'https://www.googleapis.com/auth/drive.file',
    aud: 'https://oauth2.googleapis.com/token',
    iat: Math.floor(Date.now() / 1000),
    exp: Math.floor(Date.now() / 1000) + 3600
  })).toString('base64url');
  const signed = header + '.' + payload;
  const pemKey = Buffer.from(
    SERVICE_ACCOUNT_KEY.replace(/-----BEGIN PRIVATE KEY-----/,'').replace(/-----END PRIVATE KEY-----/,'').replace(/\n/g,''), 'base64'
  );
  const signKey = await crypto.subtle.importKey('pkcs8', pemKey, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['sign']);
  const sig = Buffer.from(await crypto.subtle.sign('RSASSA-PKCS1-v1_5', signKey, Buffer.from(signed))).toString('base64url');
  const jwt = signed + '.' + sig;
  const resp = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${jwt}`
  });
  return (await resp.json()).access_token as string;
}

async function uploadPDFToDrive(pdfBuffer: Buffer, pdfName: string): Promise<string | null> {
  try {
    const token = await getDriveAccessToken();
    const boundary = 'b' + Date.now();
    const meta = JSON.stringify({ name: pdfName, parents: [DRIVE_FOLDER_ID] });
    const enc = new TextEncoder();
    const part1 = enc.encode(`--${boundary}\r\nContent-Type: application/json\r\n\r\n${meta}\r\n`);
    const part2 = enc.encode(`--${boundary}\r\nContent-Type: application/pdf\r\n\r\n`);
    const part3 = new Uint8Array(pdfBuffer);
    const part4 = enc.encode(`\r\n--${boundary}--`);
    const body = new Uint8Array(part1.byteLength + part2.byteLength + part3.byteLength + part4.byteLength);
    let off = 0;
    body.set(part1, off); off += part1.byteLength;
    body.set(part2, off); off += part2.byteLength;
    body.set(part3, off); off += part3.byteLength;
    body.set(part4, off);
    const res = await fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink',
      { method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': `multipart/related; boundary=${boundary}` }, body }
    );
    const data = await res.json();
    if (data.id) return data.webViewLink || `https://drive.google.com/file/d/${data.id}/view`;
    console.error('[Drive] Error:', JSON.stringify(data));
    return null;
  } catch (e) { console.error('[Drive] Error:', e); return null; }
}

// ==================== Email（Resend）====================
async function sendEmail(pdfBuffer: Buffer, pdfName: string, clientName: string): Promise<void> {
  try {
    const resend = new Resend(RESEND_API_KEY);
    const today = new Date().toLocaleDateString('zh-TW');
    const { error } = await resend.emails.send({
      from: EMAIL_FROM,
      to: EMAIL_TO,
      subject: `❄️ 麥好室冷氣報價單｜${clientName}｜${today}`,
      html: `
        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
          <div style="background: #1a3c5e; color: white; padding: 20px; text-align: center;">
            <h2 style="margin:0;">❄️ 麥好室冷氣報價單</h2>
            <p style="margin:5px 0 0; opacity: 0.8;">PocketPro 自動化系統</p>
          </div>
          <div style="padding: 24px; background: #f9fafb;">
            <p>您好，</p>
            <p>收到一筆新的冷氣報價單，詳細資料請參閱附件 PDF。</p>
            <table style="width: 100%; border-collapse: collapse; margin: 16px 0;">
              <tr><td style="padding: 8px 12px; color: #666;">客戶名稱</td><td style="padding: 8px 12px; font-weight: bold;">${clientName}</td></tr>
              <tr style="background: #f0f4f8;"><td style="padding: 8px 12px; color: #666;">報價日期</td><td style="padding: 8px 12px;">${today}</td></tr>
              <tr><td style="padding: 8px 12px; color: #666;">PDF 附件</td><td style="padding: 8px 12px;">📎 ${pdfName}</td></tr>
            </table>
            <p style="color: #888; font-size: 12px; margin-top: 20px;">此為系統自動發送，請勿直接回覆。如有問題請聯繫管管數位。</p>
          </div>
        </div>
      `,
      attachments: [{ filename: pdfName, content: pdfBuffer.toString('base64') }],
    });
    if (error) console.error('[Email] Resend error:', error);
    else console.log('[Email] Sent to', EMAIL_TO);
  } catch (e) { console.error('[Email] Exception:', e); }
}

// ==================== Lark ====================
async function getLarkToken(): Promise<string | null> {
  try {
    const res = await fetch('https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ app_id: LARK_APP_ID, app_secret: LARK_APP_SECRET })
    });
    const d = await res.json() as { code: number; tenant_access_token?: string };
    return d.code === 0 ? (d.tenant_access_token ?? null) : null;
  } catch { return null; }
}

async function getRecord(token: string, id: string): Promise<Record<string, unknown> | null> {
  try {
    const res = await fetch(`https://open.larksuite.com/open-apis/bitable/v1/apps/${LARK_TABLE_ID}/records/${id}`,
      { headers: { Authorization: `Bearer ${token}` } });
    const d = await res.json() as { data?: { fields?: Record<string, unknown> } };
    return d?.data?.fields ?? null;
  } catch { return null; }
}

async function updateStatus(token: string, id: string): Promise<void> {
  try {
    await fetch(`https://open.larksuite.com/open-apis/bitable/v1/apps/${LARK_TABLE_ID}/records/${id}/fields`, {
      method: 'PUT', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ fields: { '狀態': '已產出 PDF' } })
    });
  } catch { /* ignore */ }
}

// ==================== PDF ====================
async function generatePDFBuffer(fields: Record<string, unknown>): Promise<Buffer> {
  const { default: PDFDocument } = await import('pdfkit');
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'A4', margin: 20 });
    const chunks: Buffer[] = [];
    doc.on('data', (c: Buffer) => chunks.push(c));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const brand  = (fields['品牌'] as string) || '（未指定）';
    const client = String(fields['客戶名稱'] || '-');
    const project = String(fields['建案名稱'] || '-');
    const sales  = String(fields['業務人員'] || '-');
    const today  = new Date().toLocaleDateString('zh-TW');

    doc.rect(0, 0, 210, 40).fill('#1a3c5e');
    doc.fillColor('#fff').fontSize(16).font('Helvetica-Bold');
    doc.text('冷氣工程報價單', 20, 12);
    doc.fontSize(9).font('Helvetica').text(`日期：${today}`, 150, 12);
    doc.text('PocketPro 麥好室自動化系統', 20, 28);

    let y = 55;
    doc.setFillColor('#f0f4f8').rect(15, y, 180, 6).fill();
    doc.fillColor(100).fontSize(8).text('【 基本資訊 】', 20, y + 1.5);
    y += 14; doc.fillColor(0).fontSize(9);
    const rows: [string,string][] = [
      ['客戶名稱', client], ['建案名稱', project],
      ['業務人員', sales], ['電話', String(fields['電話'] || '-')]
    ];
    for (let i = 0; i < rows.length; i += 2) {
      doc.font('Helvetica-Bold').text(`${rows[i][0]}：`, 20, y);
      doc.font('Helvetica').text(rows[i][1], 55, y);
      if (rows[i+1]) {
        doc.font('Helvetica-Bold').text(`${rows[i+1][0]}：`, 120, y);
        doc.font('Helvetica').text(rows[i+1][1], 145, y);
      }
      y += 8;
    }

    y += 8;
    doc.setFillColor('#f0f4f8').rect(15, y, 180, 6).fill();
    doc.fillColor(100).fontSize(8).text('【 報價品項 】', 20, y + 1.5);
    y += 14; doc.fillColor(0).fontSize(9);

    const bCount = parseInt(String(fields['品牌冷氣數量'] || '0'));
    const bPrice = parseInt(String(fields['品牌冷氣單價'] || '0'));
    doc.text(`• ${brand} 分離式冷氣`, 20, y); y += 7;
    if (bCount > 0) { doc.text(`  數量：${bCount} 台`, 25, y); y += 7; }
    if (bCount > 0 && bPrice > 0) {
      doc.text(`  單價：$${bPrice.toLocaleString()} / 台`, 25, y); y += 7;
      doc.font('Helvetica-Bold').text(`  小計：$${(bCount * bPrice).toLocaleString()}`, 25, y); y += 7;
    }
    doc.font('Helvetica');

    const pipe = String(fields['管線材料費'] || '-');
    const pCount = parseInt(String(fields['管線數量'] || '0'));
    const pPrice = parseInt(String(fields['管線單價'] || '0'));
    y += 3; doc.text(`• 管線材料費：${pipe}`, 20, y); y += 7;
    if (pCount > 0 && pPrice > 0) {
      doc.text(`  數量：${pCount} / 單價：$${pPrice.toLocaleString()} / 小計：$${(pCount * pPrice).toLocaleString()}`, 25, y); y += 7;
    }

    const total = String(fields['總計金額'] || fields['品項總計'] || '');
    if (total) {
      y += 5;
      doc.setFillColor('#1a3c5e').rect(130, y - 5, 65, 18).fill();
      doc.fillColor('#fff').fontSize(11).font('Helvetica-Bold');
      doc.text(`合計：$${parseInt(total).toLocaleString()}`, 133, y);
      y += 20;
    }

    y += 5;
    doc.setFontSize(8).fillColor(120).font('Helvetica');
    ['※ 施工前請確認現場管線配置是否符合規範', '※ 所有費用不含稅金，如需發票請另行告知', '※ 此報價單僅供參考，實際費用以現場估價為準'].forEach(t => { doc.text(t, 20, y); y += 6; });

    y += 15;
    doc.setFontSize(9).fillColor(0);
    doc.text('業務簽名：________________', 20, y);
    doc.text('客戶確認：________________', 120, y);

    doc.end();
  });
}

// ==================== 主程式 ====================
export default async function handler(req: VercelRequest, res: VercelResponse) {
  setCors(res);
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { record_id, fields } = req.body as { record_id?: string; fields?: Record<string, unknown> };

  try {
    const token = await getLarkToken();
    if (!token) return res.status(500).json({ success: false, error: 'Cannot get Lark token' });

    let data: Record<string, unknown>;
    if (record_id) {
      data = (await getRecord(token, record_id)) || {};
      await updateStatus(token, record_id);
    } else {
      data = fields || {};
    }

    const pdfBuffer = await generatePDFBuffer(data);
    const clientName = String(data['客戶名稱'] || '新報價單');
    const pdfName = `報價單_${clientName}_${new Date().toLocaleDateString('zh-TW').replace(/\//g, '-')}.pdf`;

    const driveLink = await uploadPDFToDrive(pdfBuffer, pdfName);
    await sendEmail(pdfBuffer, pdfName, clientName);

    return res.status(200).json({ success: true, pdfName, driveLink: driveLink || null });
  } catch (err) {
    console.error('[Webhook] Error:', err);
    return res.status(500).json({ success: false, error: String(err) });
  }
}
