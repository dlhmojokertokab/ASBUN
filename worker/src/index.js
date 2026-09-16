async function getGoogleAccessToken(env) {
  const required = [
    "GOOGLE_CLIENT_ID",
    "GOOGLE_CLIENT_SECRET",
    "GOOGLE_REFRESH_TOKEN",
  ];

  for (const key of required) {
    if (!env[key]) {
      throw new Error(
        `Environment variable ${key} belum tersedia`
      );
    }
  }

  const response = await fetch(
    "https://oauth2.googleapis.com/token",
    {
      method: "POST",
      headers: {
        "Content-Type":
          "application/x-www-form-urlencoded;charset=UTF-8",
      },
      body: new URLSearchParams({
        client_id: env.GOOGLE_CLIENT_ID,
        client_secret: env.GOOGLE_CLIENT_SECRET,
        refresh_token: env.GOOGLE_REFRESH_TOKEN,
        grant_type: "refresh_token",
      }),
    }
  );

  const data = await response.json();

  if (!response.ok || !data.access_token) {
    const detail = [
      data.error,
      data.error_description,
    ]
      .filter(Boolean)
      .join(": ");

    throw new Error(
      detail ||
        "Gagal mendapatkan Google access token"
    );
  }

  return data.access_token;
}

function getSheetId(env) {
  const value =
    env.ASBUN_SHEET_ID ||
    env.GOOGLE_SPREADSHEET_ID;

  if (!value) {
    throw new Error(
      "ASBUN_SHEET_ID belum tersedia"
    );
  }

  return value;
}

function getTransitFolderId(env) {
  const value =
    env.ASBUN_TRANSIT_FOLDER_ID ||
    env.GOOGLE_TEMP_FOLDER_ID;

  if (!value) {
    throw new Error(
      "ASBUN_TRANSIT_FOLDER_ID belum tersedia"
    );
  }

  return value;
}

function cleanText(value) {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ");
}

function normalizeCompanyName(value) {
  return cleanText(value)
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(
    value
  );
}

function generatePrivateToken() {
  const bytes =
    new Uint8Array(32);

  crypto.getRandomValues(bytes);

  return Array.from(bytes)
    .map(
      byte =>
        byte
          .toString(16)
          .padStart(2, "0")
    )
    .join("");
}

function newId(prefix) {
  return `${prefix}_${crypto.randomUUID()}`;
}

function getJakartaYear() {
  return new Intl.DateTimeFormat(
    "en-US",
    {
      timeZone: "Asia/Jakarta",
      year: "numeric",
    }
  ).format(new Date());
}

function getCorsHeaders(request) {
  const origin =
    request.headers.get("Origin");

  const allowedOrigins = new Set([
    "http://localhost:5500",
    "http://127.0.0.1:5500",
  ]);

  const headers = {
    "Access-Control-Allow-Methods":
      "GET, POST, PUT, DELETE, OPTIONS",
    "Access-Control-Allow-Headers":
      "Content-Type, Authorization",
    "Vary": "Origin",
  };

  if (
    origin &&
    allowedOrigins.has(origin)
  ) {
    headers[
      "Access-Control-Allow-Origin"
    ] = origin;
  }

  return headers;
}

function jsonResponse(
  request,
  data,
  status = 200
) {
  return Response.json(
    data,
    {
      status,
      headers:
        getCorsHeaders(request),
    }
  );
}

async function readValues(
  accessToken,
  sheetId,
  range
) {
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/` +
      `${sheetId}/values/` +
      `${encodeURIComponent(range)}`,
    {
      headers: {
        Authorization:
          `Bearer ${accessToken}`,
      },
    }
  );

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      data?.error?.message ||
        `Gagal membaca ${range}`
    );
  }

  return data.values || [];
}

async function getNextRow(
  accessToken,
  sheetId,
  tabName
) {
  const rows =
    await readValues(
      accessToken,
      sheetId,
      `${tabName}!A2:A`
    );

  return rows.length + 2;
}

async function appendRow(
  accessToken,
  sheetId,
  tabName,
  row
) {
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/` +
      `${sheetId}/values/` +
      `${tabName}!A:Z:append` +
      `?valueInputOption=RAW` +
      `&insertDataOption=INSERT_ROWS`,
    {
      method: "POST",
      headers: {
        Authorization:
          `Bearer ${accessToken}`,
        "Content-Type":
          "application/json",
      },
      body: JSON.stringify({
        values: [row],
      }),
    }
  );

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      data?.error?.message ||
        `Gagal menulis ${tabName}`
    );
  }

  return data;
}

async function getNextSubmissionNumber(
  env,
  accessToken
) {
  const sheetId =
    getSheetId(env);

  const rows =
    await readValues(
      accessToken,
      sheetId,
      "CONFIG!A2:D"
    );

  const config = {};

  rows.forEach((row, index) => {
    if (!row[0]) return;

    config[row[0]] = {
      value:
        row[1] ?? "",
      rowNumber:
        index + 2,
    };
  });

  const prefix =
    config.SUBMISSION_PREFIX?.value;

  const storedYear =
    config.SEQUENCE_YEAR?.value;

  const nextSequenceRaw =
    config.NEXT_SEQUENCE?.value;

  if (!prefix) {
    throw new Error(
      "CONFIG SUBMISSION_PREFIX tidak ditemukan"
    );
  }

  if (!config.SEQUENCE_YEAR) {
    throw new Error(
      "CONFIG SEQUENCE_YEAR tidak ditemukan"
    );
  }

  if (!config.NEXT_SEQUENCE) {
    throw new Error(
      "CONFIG NEXT_SEQUENCE tidak ditemukan"
    );
  }

  const currentYear =
    getJakartaYear();

  let sequence =
    Number.parseInt(
      nextSequenceRaw,
      10
    );

  if (
    !Number.isInteger(sequence) ||
    sequence < 1
  ) {
    sequence = 1;
  }

  if (
    storedYear !== currentYear
  ) {
    sequence = 1;
  }

  const submissionNumber =
    `${prefix}-${currentYear}-` +
    String(sequence)
      .padStart(4, "0");

  const nextSequence =
    sequence + 1;

  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/` +
      `${sheetId}/values:batchUpdate`,
    {
      method: "POST",
      headers: {
        Authorization:
          `Bearer ${accessToken}`,
        "Content-Type":
          "application/json",
      },
      body: JSON.stringify({
        valueInputOption: "RAW",
        data: [
          {
            range:
              `CONFIG!B` +
              `${config.SEQUENCE_YEAR.rowNumber}`,
            values: [
              [currentYear],
            ],
          },
          {
            range:
              `CONFIG!B` +
              `${config.NEXT_SEQUENCE.rowNumber}`,
            values: [
              [
                String(
                  nextSequence
                ),
              ],
            ],
          },
        ],
      }),
    }
  );

  const data =
    await response.json();

  if (!response.ok) {
    throw new Error(
      data?.error?.message ||
        "Gagal memperbarui nomor pengajuan"
    );
  }

  return submissionNumber;
}

async function resolveCompany(
  accessToken,
  sheetId,
  companyName,
  now
) {
  const normalized =
    normalizeCompanyName(
      companyName
    );

  const rows =
    await readValues(
      accessToken,
      sheetId,
      "COMPANIES!A2:E"
    );

  const found =
    rows.find(
      row =>
        String(row[2] ?? "") ===
        normalized
    );

  if (found) {
    return {
      companyId:
        found[0],
      companyName:
        found[1],
      normalized,
      isNew: false,
      newRow: null,
    };
  }

  const companyId =
    newId("CMP");

  return {
    companyId,
    companyName,
    normalized,
    isNew: true,
    newRow: [
      companyId,
      companyName,
      normalized,
      now,
      now,
    ],
  };
}

function combineBytes(...parts) {
  const totalLength =
    parts.reduce(
      (total, part) =>
        total + part.length,
      0
    );

  const result =
    new Uint8Array(
      totalLength
    );

  let offset = 0;

  for (const part of parts) {
    result.set(
      part,
      offset
    );

    offset +=
      part.length;
  }

  return result;
}

async function uploadPdf(
  env,
  accessToken,
  pdfBuffer,
  submissionNumber,
  versionNumber = 1
) {
  const folderId =
    getTransitFolderId(env);

  const fileName =
    `${submissionNumber}_V${versionNumber}.pdf`;

  const metadata = {
    name: fileName,
    mimeType:
      "application/pdf",
    parents: [
      folderId,
    ],
  };

  const boundary =
    `asbun_${crypto.randomUUID()}`;

  const encoder =
    new TextEncoder();

  const metadataPart =
    encoder.encode(
      `--${boundary}\r\n` +
      `Content-Type: application/json; charset=UTF-8\r\n\r\n` +
      `${JSON.stringify(metadata)}\r\n` +
      `--${boundary}\r\n` +
      `Content-Type: application/pdf\r\n\r\n`
    );

  const pdfPart =
    new Uint8Array(
      pdfBuffer
    );

  const closingPart =
    encoder.encode(
      `\r\n--${boundary}--`
    );

  const multipartBody =
    combineBytes(
      metadataPart,
      pdfPart,
      closingPart
    );

  const response = await fetch(
    "https://www.googleapis.com/upload/drive/v3/files" +
      "?uploadType=multipart" +
      "&fields=id,name,mimeType,parents,size",
    {
      method: "POST",
      headers: {
        Authorization:
          `Bearer ${accessToken}`,
        "Content-Type":
          `multipart/related; boundary=${boundary}`,
      },
      body: multipartBody,
    }
  );

  const data =
    await response.json();

  if (!response.ok) {
    throw new Error(
      data?.error?.message ||
        "Gagal upload PDF ke Google Drive"
    );
  }

  return data;
}

async function deleteDriveFile(
  accessToken,
  fileId
) {
  const response = await fetch(
    `https://www.googleapis.com/drive/v3/files/${fileId}`,
    {
      method: "DELETE",
      headers: {
        Authorization:
          `Bearer ${accessToken}`,
      },
    }
  );

  return response.ok || response.status === 404;
}

async function writeSubmissionBundle(
  accessToken,
  sheetId,
  bundle
) {
  const data = [];

  if (
    bundle.company.isNew
  ) {
    const companyRow =
      await getNextRow(
        accessToken,
        sheetId,
        "COMPANIES"
      );

    data.push({
      range:
        `COMPANIES!A${companyRow}:E${companyRow}`,
      values: [
        bundle.company.newRow,
      ],
    });
  }

  const submissionRow =
    await getNextRow(
      accessToken,
      sheetId,
      "SUBMISSIONS"
    );

  const versionRow =
    await getNextRow(
      accessToken,
      sheetId,
      "VERSIONS"
    );

  const auditRow =
    await getNextRow(
      accessToken,
      sheetId,
      "AUDIT_LOG"
    );

  data.push(
    {
      range:
        `SUBMISSIONS!A${submissionRow}:N${submissionRow}`,
      values: [
        bundle.submissionRow,
      ],
    },
    {
      range:
        `VERSIONS!A${versionRow}:I${versionRow}`,
      values: [
        bundle.versionRow,
      ],
    },
    {
      range:
        `AUDIT_LOG!A${auditRow}:G${auditRow}`,
      values: [
        bundle.auditRow,
      ],
    }
  );

  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/` +
      `${sheetId}/values:batchUpdate`,
    {
      method: "POST",
      headers: {
        Authorization:
          `Bearer ${accessToken}`,
        "Content-Type":
          "application/json",
      },
      body: JSON.stringify({
        valueInputOption:
          "RAW",
        data,
      }),
    }
  );

  const result =
    await response.json();

  if (!response.ok) {
    throw new Error(
      result?.error?.message ||
        "Gagal menyimpan data pengajuan"
    );
  }

  return result;
}

function base64Url(text) {
  const bytes =
    new TextEncoder()
      .encode(text);

  let binary = "";

  for (
    let i = 0;
    i < bytes.length;
    i++
  ) {
    binary +=
      String.fromCharCode(
        bytes[i]
      );
  }

  return btoa(binary)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

async function sendInitialEmail(
  accessToken,
  {
    picEmail,
    picName,
    companyName,
    submissionNumber,
  }
) {
  const subject =
    `[ASBUN] Pengajuan ${submissionNumber} diterima`;

  const body = [
    `Yth. ${picName},`,
    ``,
    `Pengajuan proposal Anda telah diterima oleh ASBUN.`,
    ``,
    `Nomor Pengajuan: ${submissionNumber}`,
    `Perusahaan: ${companyName}`,
    `Status: MENUNGGU ASISTENSI`,
    ``,
    `Berkas saat ini menunggu proses asistensi TALING.`,
    ``,
    `Private link belum dikirim pada tahap ini. Link akan dikirim apabila berkas memerlukan perbaikan atau telah FINAL.`,
    ``,
    `ASBUN`,
    `Asistensi Berkas Usaha dan Perizinan`,
  ].join("\r\n");

  const mime = [
    `To: ${picEmail}`,
    `Subject: ${subject}`,
    `MIME-Version: 1.0`,
    `Content-Type: text/plain; charset=UTF-8`,
    `Content-Transfer-Encoding: 8bit`,
    ``,
    body,
  ].join("\r\n");

  const response = await fetch(
    "https://gmail.googleapis.com/gmail/v1/users/me/messages/send",
    {
      method: "POST",
      headers: {
        Authorization:
          `Bearer ${accessToken}`,
        "Content-Type":
          "application/json",
      },
      body: JSON.stringify({
        raw:
          base64Url(mime),
      }),
    }
  );

  const data =
    await response.json();

  if (!response.ok) {
    throw new Error(
      data?.error?.message ||
        "Gagal mengirim email"
    );
  }

  return data;
}

async function parseSubmission(
  request
) {
  const contentType =
    request.headers.get(
      "Content-Type"
    ) || "";

  if (
    !contentType
      .toLowerCase()
      .includes(
        "multipart/form-data"
      )
  ) {
    throw new Error(
      "Content-Type harus multipart/form-data"
    );
  }

  const form =
    await request.formData();

  const companyName =
    cleanText(
      form.get(
        "companyName"
      )
    );

  const picName =
    cleanText(
      form.get(
        "picName"
      )
    );

  const picEmail =
    cleanText(
      form.get(
        "picEmail"
      )
    ).toLowerCase();

  const proposal =
    form.get(
      "proposal"
    );

  if (!companyName) {
    throw new Error(
      "Nama perusahaan wajib diisi"
    );
  }

  if (!picName) {
    throw new Error(
      "Nama PIC wajib diisi"
    );
  }

  if (
    !picEmail ||
    !isValidEmail(picEmail)
  ) {
    throw new Error(
      "Email PIC tidak valid"
    );
  }

  const isFile =
    proposal &&
    typeof proposal === "object" &&
    typeof proposal.arrayBuffer ===
      "function";

  if (!isFile) {
    throw new Error(
      "Proposal PDF wajib diunggah"
    );
  }

  if (
    proposal.size <= 0
  ) {
    throw new Error(
      "File proposal kosong"
    );
  }

  const pdfBuffer =
    await proposal.arrayBuffer();

  const signature =
    new TextDecoder()
      .decode(
        new Uint8Array(
          pdfBuffer
        ).slice(0, 5)
      );

  if (
    signature !== "%PDF-"
  ) {
    throw new Error(
      "File proposal bukan PDF yang valid"
    );
  }

  return {
    companyName,
    picName,
    picEmail,
    proposalName:
      proposal.name ||
      "proposal.pdf",
    pdfBuffer,
  };
}

async function createSubmission(
  env,
  parsed
) {
  const accessToken =
    await getGoogleAccessToken(
      env
    );

  const sheetId =
    getSheetId(env);

  const now =
    new Date().toISOString();

  const company =
    await resolveCompany(
      accessToken,
      sheetId,
      parsed.companyName,
      now
    );

  const submissionNumber =
    await getNextSubmissionNumber(
      env,
      accessToken
    );

  const submissionId =
    newId("SUB");

  const versionId =
    newId("VER");

  const auditId =
    newId("LOG");

  const privateToken =
    generatePrivateToken();

  let uploadedFile = null;

  try {
    uploadedFile =
      await uploadPdf(
        env,
        accessToken,
        parsed.pdfBuffer,
        submissionNumber
      );

    const submissionRow = [
      submissionId,
      submissionNumber,
      company.companyId,
      parsed.picName,
      parsed.picEmail,
      privateToken,
      "MENUNGGU_ASISTENSI",
      uploadedFile.id,
      uploadedFile.name,
      "1",
      now,
      now,
      "",
      "",
    ];

    const versionRow = [
      versionId,
      submissionId,
      "1",
      uploadedFile.id,
      uploadedFile.name,
      String(
        uploadedFile.size ??
        parsed.pdfBuffer.byteLength
      ),
      "PIC",
      now,
      "",
    ];

    const auditRow = [
      auditId,
      submissionId,
      "SUBMISSION_CREATED",
      "PIC",
      parsed.picName,
      JSON.stringify({
        submissionNumber,
        originalFileName:
          parsed.proposalName,
        version: 1,
      }),
      now,
    ];

    await writeSubmissionBundle(
      accessToken,
      sheetId,
      {
        company,
        submissionRow,
        versionRow,
        auditRow,
      }
    );
  } catch (error) {
    if (
      uploadedFile?.id
    ) {
      await deleteDriveFile(
        accessToken,
        uploadedFile.id
      ).catch(
        () => false
      );
    }

    throw error;
  }

  let emailSent = false;
  let emailError = null;

  try {
    await sendInitialEmail(
      accessToken,
      {
        picEmail:
          parsed.picEmail,
        picName:
          parsed.picName,
        companyName:
          company.companyName,
        submissionNumber,
      }
    );

    emailSent = true;

    await appendRow(
      accessToken,
      sheetId,
      "AUDIT_LOG",
      [
        newId("LOG"),
        submissionId,
        "EMAIL_INITIAL_SENT",
        "SYSTEM",
        "ASBUN",
        parsed.picEmail,
        new Date().toISOString(),
      ]
    );
  } catch (error) {
    emailError =
      error.message;

    await appendRow(
      accessToken,
      sheetId,
      "AUDIT_LOG",
      [
        newId("LOG"),
        submissionId,
        "EMAIL_INITIAL_FAILED",
        "SYSTEM",
        "ASBUN",
        emailError,
        new Date().toISOString(),
      ]
    ).catch(
      () => null
    );
  }

  return {
    submissionNumber,
    status:
      "MENUNGGU_ASISTENSI",
    companyName:
      company.companyName,
    companyCreated:
      company.isNew,
    activeVersion: 1,
    emailSent,
    emailError,
  };
}


function httpError(status, message) {
  const error = new Error(message);
  error.status = status;
  return error;
}

async function findSubmissionByNumber(
  accessToken,
  sheetId,
  submissionNumber
) {
  const rows = await readValues(
    accessToken,
    sheetId,
    "SUBMISSIONS!A2:N"
  );

  const index = rows.findIndex(
    row => String(row[1] ?? "") === submissionNumber
  );

  if (index === -1) {
    return null;
  }

  return {
    rowNumber: index + 2,
    row: rows[index],
  };
}

async function sendNeedsRevisionEmail(
  env,
  accessToken,
  {
    picEmail,
    picName,
    companyName,
    submissionNumber,
    note,
    privateToken,
  }
) {
  const frontendUrl =
    String(env.ASBUN_FRONTEND_URL || "")
      .trim()
      .replace(/\/+$/, "");

  if (!frontendUrl) {
    throw new Error(
      "ASBUN_FRONTEND_URL belum tersedia"
    );
  }

  if (!privateToken) {
    throw new Error(
      "Private token pengajuan tidak tersedia"
    );
  }

  const privateLink =
    frontendUrl +
    "/?t=" +
    encodeURIComponent(privateToken);

  const subject =
    "[ASBUN] Perbaikan diperlukan - " +
    submissionNumber;

  const body = [
    "Yth. " + picName + ",",
    "",
    "Hasil asistensi TALING menunjukkan bahwa proposal masih memerlukan perbaikan.",
    "",
    "Nomor Pengajuan: " + submissionNumber,
    "Perusahaan: " + companyName,
    "Status: PERLU PERBAIKAN",
    "",
    "Catatan asistensi:",
    note,
    "",
    "Silakan buka private link berikut untuk melihat pengajuan dan mengunggah versi perbaikan:",
    privateLink,
    "",
    "Private link ini khusus untuk pengajuan Anda. Mohon tidak membagikannya kepada pihak lain.",
    "",
    "ASBUN",
    "Asistensi Berkas Usaha dan Perizinan",
  ].join("\r\n");

  const mime = [
    "To: " + picEmail,
    "Subject: " + subject,
    "MIME-Version: 1.0",
    "Content-Type: text/plain; charset=UTF-8",
    "Content-Transfer-Encoding: 8bit",
    "",
    body,
  ].join("\r\n");

  const response = await fetch(
    "https://gmail.googleapis.com/gmail/v1/users/me/messages/send",
    {
      method: "POST",
      headers: {
        Authorization: "Bearer " + accessToken,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        raw: base64Url(mime),
      }),
    }
  );

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      data?.error?.message ||
      "Gagal mengirim email perbaikan"
    );
  }

  return data;
}

async function markNeedsRevision(
  env,
  {
    submissionNumber,
    note,
    admin,
  }
) {
  submissionNumber =
    cleanText(submissionNumber);

  note =
    cleanText(note);

  const createdBy =
    cleanText(
      admin?.displayName
    ) ||
    cleanText(
      admin?.username
    );

  if (
    !admin?.adminId ||
    !createdBy
  ) {
    throw httpError(
      401,
      "Identitas admin tidak valid"
    );
  }

  if (!submissionNumber) {
    throw httpError(
      400,
      "Nomor pengajuan wajib diisi"
    );
  }

  if (!note) {
    throw httpError(
      400,
      "Catatan asistensi wajib diisi"
    );
  }

  const accessToken =
    await getGoogleAccessToken(env);

  const sheetId =
    getSheetId(env);

  const found =
    await findSubmissionByNumber(
      accessToken,
      sheetId,
      submissionNumber
    );

  if (!found) {
    throw httpError(
      404,
      "Pengajuan tidak ditemukan"
    );
  }

  const row =
    found.row;

  const submissionId = row[0];
  const companyId = row[2];
  const picName = row[3];
  const picEmail = row[4];
  const privateToken = row[5];
  const currentStatus = row[6];
  const activeVersion = row[9] || "1";

  if (currentStatus !== "MENUNGGU_ASISTENSI") {
    throw httpError(
      409,
      "Status pengajuan saat ini " +
      currentStatus +
      ", sehingga tidak dapat ditandai PERLU_PERBAIKAN"
    );
  }

  if (!privateToken) {
    throw new Error(
      "Private token pengajuan tidak ditemukan"
    );
  }

  const companies =
    await readValues(
      accessToken,
      sheetId,
      "COMPANIES!A2:E"
    );

  const company =
    companies.find(
      item => item[0] === companyId
    );

  const companyName =
    company?.[1] || companyId;

  const now =
    new Date().toISOString();

  const noteId =
    newId("NOTE");

  const auditId =
    newId("LOG");

  const noteRow =
    await getNextRow(
      accessToken,
      sheetId,
      "ASSISTANCE_NOTES"
    );

  const auditRow =
    await getNextRow(
      accessToken,
      sheetId,
      "AUDIT_LOG"
    );

  const writeResponse = await fetch(
    "https://sheets.googleapis.com/v4/spreadsheets/" +
      sheetId +
      "/values:batchUpdate",
    {
      method: "POST",
      headers: {
        Authorization: "Bearer " + accessToken,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        valueInputOption: "RAW",
        data: [
          {
            range:
              "SUBMISSIONS!G" +
              found.rowNumber,
            values: [
              ["PERLU_PERBAIKAN"],
            ],
          },
          {
            range:
              "SUBMISSIONS!L" +
              found.rowNumber,
            values: [
              [now],
            ],
          },
          {
            range:
              "ASSISTANCE_NOTES!A" +
              noteRow +
              ":F" +
              noteRow,
            values: [
              [
                noteId,
                submissionId,
                String(activeVersion),
                note,
                createdBy,
                now,
              ],
            ],
          },
          {
            range:
              "AUDIT_LOG!A" +
              auditRow +
              ":G" +
              auditRow,
            values: [
              [
                auditId,
                submissionId,
                "MARKED_NEEDS_REVISION",
                "ADMIN",
                createdBy,
                JSON.stringify({
                  submissionNumber,
                  version:
                    Number(activeVersion),
                  noteId,
                  adminId:
                    admin.adminId,
                  adminUsername:
                    admin.username,
                }),
                now,
              ],
            ],
          },
        ],
      }),
    }
  );

  const writeData =
    await writeResponse.json();

  if (!writeResponse.ok) {
    throw new Error(
      writeData?.error?.message ||
      "Gagal menyimpan keputusan asistensi"
    );
  }

  let emailSent = false;
  let emailError = null;

  try {
    await sendNeedsRevisionEmail(
      env,
      accessToken,
      {
        picEmail,
        picName,
        companyName,
        submissionNumber,
        note,
        privateToken,
      }
    );

    emailSent = true;

    await appendRow(
      accessToken,
      sheetId,
      "AUDIT_LOG",
      [
        newId("LOG"),
        submissionId,
        "EMAIL_REVISION_SENT",
        "SYSTEM",
        "ASBUN",
        picEmail,
        new Date().toISOString(),
      ]
    );
  } catch (error) {
    emailError = error.message;

    await appendRow(
      accessToken,
      sheetId,
      "AUDIT_LOG",
      [
        newId("LOG"),
        submissionId,
        "EMAIL_REVISION_FAILED",
        "SYSTEM",
        "ASBUN",
        emailError,
        new Date().toISOString(),
      ]
    ).catch(() => null);
  }

  return {
    submissionNumber,
    status: "PERLU_PERBAIKAN",
    activeVersion:
      Number(activeVersion),
    note,
    emailSent,
    emailError,
  };
}


async function parseRevisionUpload(request) {
  const contentType =
    request.headers.get("Content-Type") || "";

  if (
    !contentType
      .toLowerCase()
      .includes("multipart/form-data")
  ) {
    throw httpError(
      415,
      "Content-Type harus multipart/form-data"
    );
  }

  const form =
    await request.formData();

  const privateToken =
    cleanText(
      form.get("token")
    );

  const proposal =
    form.get("proposal");

  if (!privateToken) {
    throw httpError(
      400,
      "Private token wajib disertakan"
    );
  }

  const isFile =
    proposal &&
    typeof proposal === "object" &&
    typeof proposal.arrayBuffer ===
      "function";

  if (!isFile) {
    throw httpError(
      400,
      "Proposal PDF wajib diunggah"
    );
  }

  if (proposal.size <= 0) {
    throw httpError(
      400,
      "File proposal kosong"
    );
  }

  const pdfBuffer =
    await proposal.arrayBuffer();

  const signature =
    new TextDecoder()
      .decode(
        new Uint8Array(
          pdfBuffer
        ).slice(0, 5)
      );

  if (signature !== "%PDF-") {
    throw httpError(
      400,
      "File proposal bukan PDF yang valid"
    );
  }

  return {
    privateToken,
    pdfBuffer,
    proposalName:
      proposal.name ||
      "proposal.pdf",
  };
}

async function findSubmissionByPrivateToken(
  accessToken,
  sheetId,
  privateToken
) {
  const rows =
    await readValues(
      accessToken,
      sheetId,
      "SUBMISSIONS!A2:N"
    );

  const index =
    rows.findIndex(
      row =>
        String(row[5] ?? "") ===
        privateToken
    );

  if (index === -1) {
    return null;
  }

  return {
    rowNumber: index + 2,
    row: rows[index],
  };
}

async function uploadRevision(
  env,
  parsed
) {
  const accessToken =
    await getGoogleAccessToken(env);

  const sheetId =
    getSheetId(env);

  const found =
    await findSubmissionByPrivateToken(
      accessToken,
      sheetId,
      parsed.privateToken
    );

  if (!found) {
    throw httpError(
      404,
      "Private link tidak valid atau pengajuan tidak ditemukan"
    );
  }

  const row =
    found.row;

  const submissionId =
    row[0];

  const submissionNumber =
    row[1];

  const currentStatus =
    row[6];

  const oldFileId =
    row[7];

  const oldFileName =
    row[8];

  let currentVersion =
    Number.parseInt(
      row[9] || "1",
      10
    );

  if (
    !Number.isInteger(currentVersion) ||
    currentVersion < 1
  ) {
    currentVersion = 1;
  }

  if (
    currentStatus !==
    "PERLU_PERBAIKAN"
  ) {
    throw httpError(
      409,
      "Pengajuan tidak sedang berstatus PERLU_PERBAIKAN"
    );
  }

  const versions =
    await readValues(
      accessToken,
      sheetId,
      "VERSIONS!A2:I"
    );

  const oldVersionIndex =
    versions.findIndex(
      version =>
        version[1] === submissionId &&
        Number(version[2]) ===
          currentVersion
    );

  if (oldVersionIndex === -1) {
    throw new Error(
      "Metadata versi aktif lama tidak ditemukan"
    );
  }

  const oldVersionRowNumber =
    oldVersionIndex + 2;

  const newVersion =
    currentVersion + 1;

  const now =
    new Date().toISOString();

  let uploadedFile = null;

  try {
    // 1. FILE BARU MASUK DULU
    uploadedFile =
      await uploadPdf(
        env,
        accessToken,
        parsed.pdfBuffer,
        submissionNumber,
        newVersion
      );

    // 2. BARU DATABASE DIPINDAH KE VERSI BARU
    const newVersionRow =
      await getNextRow(
        accessToken,
        sheetId,
        "VERSIONS"
      );

    const auditRow =
      await getNextRow(
        accessToken,
        sheetId,
        "AUDIT_LOG"
      );

    const versionId =
      newId("VER");

    const auditId =
      newId("LOG");

    const updateResponse =
      await fetch(
        "https://sheets.googleapis.com/v4/spreadsheets/" +
          sheetId +
          "/values:batchUpdate",
        {
          method: "POST",
          headers: {
            Authorization:
              "Bearer " +
              accessToken,
            "Content-Type":
              "application/json",
          },
          body: JSON.stringify({
            valueInputOption:
              "RAW",
            data: [
              {
                range:
                  "SUBMISSIONS!G" +
                  found.rowNumber +
                  ":L" +
                  found.rowNumber,
                values: [
                  [
                    "MENUNGGU_ASISTENSI",
                    uploadedFile.id,
                    uploadedFile.name,
                    String(newVersion),
                    row[10] || "",
                    now,
                  ],
                ],
              },
              {
                range:
                  "VERSIONS!A" +
                  newVersionRow +
                  ":I" +
                  newVersionRow,
                values: [
                  [
                    versionId,
                    submissionId,
                    String(newVersion),
                    uploadedFile.id,
                    uploadedFile.name,
                    String(
                      uploadedFile.size ??
                      parsed.pdfBuffer.byteLength
                    ),
                    "PIC",
                    now,
                    "",
                  ],
                ],
              },
              {
                range:
                  "AUDIT_LOG!A" +
                  auditRow +
                  ":G" +
                  auditRow,
                values: [
                  [
                    auditId,
                    submissionId,
                    "REVISION_UPLOADED",
                    "PIC",
                    row[3] || "PIC",
                    JSON.stringify({
                      submissionNumber,
                      fromVersion:
                        currentVersion,
                      toVersion:
                        newVersion,
                      originalFileName:
                        parsed.proposalName,
                    }),
                    now,
                  ],
                ],
              },
            ],
          }),
        }
      );

    const updateData =
      await updateResponse.json();

    if (!updateResponse.ok) {
      throw new Error(
        updateData?.error?.message ||
        "Gagal memperbarui data versi"
      );
    }
  } catch (error) {
    // Kalau DB gagal setelah V2 masuk,
    // jangan biarkan file V2 yatim.
    if (uploadedFile?.id) {
      await deleteDriveFile(
        accessToken,
        uploadedFile.id
      ).catch(() => false);
    }

    throw error;
  }

  // 3. SETELAH V2 + DATABASE AMAN,
  // BARU FILE FISIK V1 DIHAPUS.
  let oldFileDeleted = false;

  if (oldFileId) {
    oldFileDeleted =
      await deleteDriveFile(
        accessToken,
        oldFileId
      );
  } else {
    oldFileDeleted = true;
  }

  if (oldFileDeleted) {
    const deletedAt =
      new Date().toISOString();

    const markResponse =
      await fetch(
        "https://sheets.googleapis.com/v4/spreadsheets/" +
          sheetId +
          "/values/VERSIONS!I" +
          oldVersionRowNumber +
          "?valueInputOption=RAW",
        {
          method: "PUT",
          headers: {
            Authorization:
              "Bearer " +
              accessToken,
            "Content-Type":
              "application/json",
          },
          body: JSON.stringify({
            values: [
              [deletedAt],
            ],
          }),
        }
      );

    if (!markResponse.ok) {
      throw new Error(
        "File lama sudah dihapus tetapi FILE_DELETED_AT gagal diperbarui"
      );
    }

    await appendRow(
      accessToken,
      sheetId,
      "AUDIT_LOG",
      [
        newId("LOG"),
        submissionId,
        "OLD_VERSION_FILE_DELETED",
        "SYSTEM",
        "ASBUN",
        JSON.stringify({
          version:
            currentVersion,
          fileId:
            oldFileId,
          fileName:
            oldFileName,
        }),
        deletedAt,
      ]
    );
  } else {
    await appendRow(
      accessToken,
      sheetId,
      "AUDIT_LOG",
      [
        newId("LOG"),
        submissionId,
        "OLD_VERSION_FILE_DELETE_FAILED",
        "SYSTEM",
        "ASBUN",
        JSON.stringify({
          version:
            currentVersion,
          fileId:
            oldFileId,
        }),
        new Date().toISOString(),
      ]
    ).catch(() => null);
  }

  return {
    submissionNumber,
    status:
      "MENUNGGU_ASISTENSI",
    previousVersion:
      currentVersion,
    activeVersion:
      newVersion,
    activeFileName:
      uploadedFile.name,
    oldFileDeleted,
  };
}


async function sendFinalEmail(
  accessToken,
  {
    picEmail,
    picName,
    companyName,
    submissionNumber,
  }
) {
  const subject =
    "[ASBUN] Asistensi selesai - " +
    submissionNumber;

  const body = [
    "Yth. " + picName + ",",
    "",
    "Proses asistensi proposal Anda telah selesai.",
    "",
    "Nomor Pengajuan: " + submissionNumber,
    "Perusahaan: " + companyName,
    "Status: FINAL",
    "",
    "Dokumen Anda telah diterima oleh TALING.",
    "Silakan menunggu proses selanjutnya.",
    "",
    "Terima kasih.",
    "",
    "ASBUN",
    "Asistensi Berkas Usaha dan Perizinan",
  ].join("\r\n");

  const mime = [
    "To: " + picEmail,
    "Subject: " + subject,
    "MIME-Version: 1.0",
    "Content-Type: text/plain; charset=UTF-8",
    "Content-Transfer-Encoding: 8bit",
    "",
    body,
  ].join("\r\n");

  const response = await fetch(
    "https://gmail.googleapis.com/gmail/v1/users/me/messages/send",
    {
      method: "POST",
      headers: {
        Authorization: "Bearer " + accessToken,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        raw: base64Url(mime),
      }),
    }
  );

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      data?.error?.message ||
      "Gagal mengirim email FINAL"
    );
  }

  return data;
}

async function markSubmissionFinal(
  env,
  {
    submissionNumber,
    admin,
  }
) {
  submissionNumber =
    cleanText(submissionNumber);

  const createdBy =
    cleanText(
      admin?.displayName
    ) ||
    cleanText(
      admin?.username
    );

  if (
    !admin?.adminId ||
    !createdBy
  ) {
    throw httpError(
      401,
      "Identitas admin tidak valid"
    );
  }

  const accessToken =
    await getGoogleAccessToken(env);

  const sheetId =
    getSheetId(env);

  const found =
    await findSubmissionByNumber(
      accessToken,
      sheetId,
      submissionNumber
    );

  if (!found) {
    throw httpError(
      404,
      "Pengajuan tidak ditemukan"
    );
  }

  const row = found.row;

  const submissionId = row[0];
  const companyId = row[2];
  const picName = row[3];
  const picEmail = row[4];
  const privateToken = row[5];
  const currentStatus = row[6];
  const activeVersion = row[9] || "1";

  if (
    currentStatus !==
    "MENUNGGU_ASISTENSI"
  ) {
    throw httpError(
      409,
      "Status pengajuan saat ini " +
      currentStatus +
      ", sehingga tidak dapat ditandai FINAL"
    );
  }

  if (!row[7]) {
    throw new Error(
      "File aktif pengajuan tidak ditemukan"
    );
  }

  const companies =
    await readValues(
      accessToken,
      sheetId,
      "COMPANIES!A2:E"
    );

  const company =
    companies.find(
      item => item[0] === companyId
    );

  const companyName =
    company?.[1] || companyId;

  const now =
    new Date().toISOString();

  const auditRow =
    await getNextRow(
      accessToken,
      sheetId,
      "AUDIT_LOG"
    );

  const response = await fetch(
    "https://sheets.googleapis.com/v4/spreadsheets/" +
      sheetId +
      "/values:batchUpdate",
    {
      method: "POST",
      headers: {
        Authorization:
          "Bearer " + accessToken,
        "Content-Type":
          "application/json",
      },
      body: JSON.stringify({
        valueInputOption: "RAW",
        data: [
          {
            range:
              "SUBMISSIONS!G" +
              found.rowNumber,
            values: [
              ["FINAL"],
            ],
          },
          {
            range:
              "SUBMISSIONS!L" +
              found.rowNumber +
              ":M" +
              found.rowNumber,
            values: [
              [
                now,
                now,
              ],
            ],
          },
          {
            range:
              "AUDIT_LOG!A" +
              auditRow +
              ":G" +
              auditRow,
            values: [
              [
                newId("LOG"),
                submissionId,
                "MARKED_FINAL",
                "ADMIN",
                createdBy,
                JSON.stringify({
                  submissionNumber:
                    submissionNumber,
                  version:
                    Number(activeVersion),
                  adminId:
                    admin.adminId,
                  adminUsername:
                    admin.username,
                }),
                now,
              ],
            ],
          },
        ],
      }),
    }
  );

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      data?.error?.message ||
      "Gagal menyimpan status FINAL"
    );
  }

  let emailSent = false;
  let emailError = null;

  try {
    await sendFinalEmail(
      accessToken,
      {
        picEmail,
        picName,
        companyName,
        submissionNumber,
      }
    );

    emailSent = true;

    await appendRow(
      accessToken,
      sheetId,
      "AUDIT_LOG",
      [
        newId("LOG"),
        submissionId,
        "EMAIL_FINAL_SENT",
        "SYSTEM",
        "ASBUN",
        picEmail,
        new Date().toISOString(),
      ]
    );
  } catch (error) {
    emailError = error.message;

    await appendRow(
      accessToken,
      sheetId,
      "AUDIT_LOG",
      [
        newId("LOG"),
        submissionId,
        "EMAIL_FINAL_FAILED",
        "SYSTEM",
        "ASBUN",
        emailError,
        new Date().toISOString(),
      ]
    ).catch(() => null);
  }

  return {
    submissionNumber,
    status: "FINAL",
    activeVersion:
      Number(activeVersion),
    emailSent,
    emailError,
  };
}

async function completeSubmission(
  env,
  {
    submissionNumber,
    admin,
  }
) {
  submissionNumber =
    cleanText(submissionNumber);

  const completedBy =
    cleanText(
      admin?.displayName
    ) ||
    cleanText(
      admin?.username
    );

  if (
    !admin?.adminId ||
    !completedBy
  ) {
    throw httpError(
      401,
      "Identitas admin tidak valid"
    );
  }

  const accessToken =
    await getGoogleAccessToken(env);

  const sheetId =
    getSheetId(env);

  const found =
    await findSubmissionByNumber(
      accessToken,
      sheetId,
      submissionNumber
    );

  if (!found) {
    throw httpError(
      404,
      "Pengajuan tidak ditemukan"
    );
  }

  const row = found.row;

  const submissionId = row[0];
  const currentStatus = row[6];
  const activeFileId = row[7];
  const activeFileName = row[8];

  let activeVersion =
    Number.parseInt(
      row[9] || "1",
      10
    );

  if (
    !Number.isInteger(activeVersion) ||
    activeVersion < 1
  ) {
    activeVersion = 1;
  }

  if (currentStatus !== "FINAL") {
    throw httpError(
      409,
      "Hanya pengajuan berstatus FINAL yang dapat diselesaikan"
    );
  }

  const versions =
    await readValues(
      accessToken,
      sheetId,
      "VERSIONS!A2:I"
    );

  const versionIndex =
    versions.findIndex(
      item =>
        item[1] === submissionId &&
        Number(item[2]) === activeVersion
    );

  if (versionIndex === -1) {
    throw new Error(
      "Metadata versi aktif tidak ditemukan"
    );
  }

  if (activeFileId) {
    const deleted =
      await deleteDriveFile(
        accessToken,
        activeFileId
      );

    if (!deleted) {
      throw new Error(
        "Gagal menghapus file FINAL dari Google Drive"
      );
    }
  }

  const completedAt =
    new Date().toISOString();

  const versionRow =
    versionIndex + 2;

  const auditRow =
    await getNextRow(
      accessToken,
      sheetId,
      "AUDIT_LOG"
    );

  const response = await fetch(
    "https://sheets.googleapis.com/v4/spreadsheets/" +
      sheetId +
      "/values:batchUpdate",
    {
      method: "POST",
      headers: {
        Authorization:
          "Bearer " + accessToken,
        "Content-Type":
          "application/json",
      },
      body: JSON.stringify({
        valueInputOption: "RAW",
        data: [
          {
            range:
              "SUBMISSIONS!G" +
              found.rowNumber +
              ":N" +
              found.rowNumber,
            values: [
              [
                "SELESAI",
                "",
                "",
                String(activeVersion),
                row[10] || "",
                completedAt,
                row[12] || "",
                completedAt,
              ],
            ],
          },
          {
            range:
              "VERSIONS!I" +
              versionRow,
            values: [
              [completedAt],
            ],
          },
          {
            range:
              "AUDIT_LOG!A" +
              auditRow +
              ":G" +
              auditRow,
            values: [
              [
                newId("LOG"),
                submissionId,
                "SUBMISSION_COMPLETED_FILE_DELETED",
                "ADMIN",
                completedBy,
                JSON.stringify({
                  submissionNumber:
                    submissionNumber,
                  version:
                    activeVersion,
                  fileId:
                    activeFileId || null,
                  fileName:
                    activeFileName || null,
                  adminId:
                    admin.adminId,
                  adminUsername:
                    admin.username,
                }),
                completedAt,
              ],
            ],
          },
        ],
      }),
    }
  );

  const data =
    await response.json();

  if (!response.ok) {
    throw new Error(
      data?.error?.message ||
      "File sudah dihapus tetapi database gagal diperbarui"
    );
  }

  return {
    submissionNumber,
    status: "SELESAI",
    activeVersion,
    fileDeleted: true,
    completedAt,
  };
}


async function getPrivateSubmissionView(
  env,
  privateToken
) {
  privateToken =
    cleanText(privateToken);

  if (!privateToken) {
    throw httpError(
      400,
      "Private token wajib disertakan"
    );
  }

  const accessToken =
    await getGoogleAccessToken(env);

  const sheetId =
    getSheetId(env);

  const found =
    await findSubmissionByPrivateToken(
      accessToken,
      sheetId,
      privateToken
    );

  if (!found) {
    throw httpError(
      404,
      "Private link tidak valid atau pengajuan tidak ditemukan"
    );
  }

  const row =
    found.row;

  const submissionId =
    row[0];

  const submissionNumber =
    row[1];

  const companyId =
    row[2];

  const picName =
    row[3];

  const status =
    row[6];

  const activeFileName =
    row[8] || null;

  let activeVersion =
    Number.parseInt(
      row[9] || "1",
      10
    );

  if (
    !Number.isInteger(activeVersion) ||
    activeVersion < 1
  ) {
    activeVersion = 1;
  }

  const finalAt =
    row[12] || null;

  const completedAt =
    row[13] || null;

  const companies =
    await readValues(
      accessToken,
      sheetId,
      "COMPANIES!A2:E"
    );

  const company =
    companies.find(
      item =>
        item[0] === companyId
    );

  const companyName =
    company?.[1] || companyId;

  const notes =
    await readValues(
      accessToken,
      sheetId,
      "ASSISTANCE_NOTES!A2:F"
    );

  const ownNotes =
    notes
      .filter(
        item =>
          item[1] === submissionId
      )
      .map(
        item => ({
          versionNumber:
            Number(item[2] || 0),
          note:
            item[3] || "",
          createdBy:
            item[4] || "TALING",
          createdAt:
            item[5] || null,
        })
      );

  return {
    submissionNumber,
    companyName,
    picName,
    status,
    activeVersion,
    activeFileName,
    fileAvailable:
      Boolean(row[7]) &&
      status !== "FINAL" &&
      status !== "SELESAI",
    canUploadRevision:
      status ===
      "PERLU_PERBAIKAN",
    isFinal:
      status === "FINAL",
    isCompleted:
      status === "SELESAI",
    finalAt,
    completedAt,
    notes:
      ownNotes,
  };
}


async function getPrivateActiveFile(
  env,
  privateToken
) {
  privateToken =
    cleanText(privateToken);

  if (!privateToken) {
    throw httpError(
      400,
      "Private token wajib disertakan"
    );
  }

  const accessToken =
    await getGoogleAccessToken(env);

  const sheetId =
    getSheetId(env);

  const found =
    await findSubmissionByPrivateToken(
      accessToken,
      sheetId,
      privateToken
    );

  if (!found) {
    throw httpError(
      404,
      "Private link tidak valid atau pengajuan tidak ditemukan"
    );
  }

  const row =
    found.row;

  const status =
    row[6];

  const fileId =
    row[7];

  const fileName =
    row[8] ||
    "proposal.pdf";

  if (status === "FINAL") {
    throw httpError(
      403,
      "Proses asistensi telah selesai. Berkas final dikelola oleh TALING."
    );
  }

  if (status === "SELESAI") {
    throw httpError(
      410,
      "Berkas sudah dihapus karena proses pengajuan telah selesai"
    );
  }

  if (!fileId) {
    throw httpError(
      404,
      "Berkas aktif tidak tersedia"
    );
  }

  const response =
    await fetch(
      "https://www.googleapis.com/drive/v3/files/" +
        encodeURIComponent(fileId) +
        "?alt=media",
      {
        headers: {
          Authorization:
            "Bearer " +
            accessToken,
        },
      }
    );

  if (response.status === 404) {
    throw httpError(
      404,
      "Berkas aktif tidak ditemukan"
    );
  }

  if (!response.ok) {
    throw new Error(
      "Gagal membaca berkas dari Google Drive"
    );
  }

  const buffer =
    await response.arrayBuffer();

  if (!buffer.byteLength) {
    throw new Error(
      "Berkas aktif kosong"
    );
  }

  const signature =
    new TextDecoder()
      .decode(
        new Uint8Array(
          buffer
        ).slice(0, 5)
      );

  if (signature !== "%PDF-") {
    throw new Error(
      "Berkas aktif bukan PDF yang valid"
    );
  }

  return {
    buffer,
    fileName,
  };
}


function base64UrlToBytes(value) {
  const normalized =
    String(value || "")
      .replace(/-/g, "+")
      .replace(/_/g, "/");

  const padded =
    normalized +
    "=".repeat(
      (4 - (normalized.length % 4)) % 4
    );

  const binary =
    atob(padded);

  return Uint8Array.from(
    binary,
    (char) =>
      char.charCodeAt(0)
  );
}

function bytesToBase64Url(bytes) {
  let binary = "";

  for (const byte of bytes) {
    binary +=
      String.fromCharCode(byte);
  }

  return btoa(binary)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

async function verifyAdminPassword(
  password,
  storedHash,
  storedSalt
) {
  const parts =
    String(storedHash || "")
      .split("$");

  if (
    parts.length !== 3 ||
    parts[0] !== "pbkdf2-sha256"
  ) {
    return false;
  }

  const iterations =
    Number.parseInt(
      parts[1],
      10
    );

  if (
    !Number.isInteger(iterations) ||
    iterations < 1
  ) {
    return false;
  }

  const expectedHash =
    base64UrlToBytes(
      parts[2]
    );

  const salt =
    base64UrlToBytes(
      storedSalt
    );

  const keyMaterial =
    await crypto.subtle.importKey(
      "raw",
      new TextEncoder().encode(
        String(password || "")
      ),
      "PBKDF2",
      false,
      ["deriveBits"]
    );

  const derivedHash =
    new Uint8Array(
      await crypto.subtle.deriveBits(
        {
          name: "PBKDF2",
          hash: "SHA-256",
          salt,
          iterations,
        },
        keyMaterial,
        expectedHash.length * 8
      )
    );

  if (
    derivedHash.length !==
    expectedHash.length
  ) {
    return false;
  }

  let difference = 0;

  for (
    let i = 0;
    i < derivedHash.length;
    i++
  ) {
    difference |=
      derivedHash[i] ^
      expectedHash[i];
  }

  return difference === 0;
}

async function getAdminSessionKey(env) {
  if (!env.ASBUN_SESSION_SECRET) {
    throw new Error(
      "ASBUN_SESSION_SECRET belum tersedia"
    );
  }

  return crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(
      env.ASBUN_SESSION_SECRET
    ),
    {
      name: "HMAC",
      hash: "SHA-256",
    },
    false,
    ["sign", "verify"]
  );
}

async function createAdminSession(
  env,
  admin
) {
  const now =
    Math.floor(
      Date.now() / 1000
    );

  const payload = {
    adminId:
      admin.adminId,
    username:
      admin.username,
    iat:
      now,
    exp:
      now + 8 * 60 * 60,
  };

  const encodedPayload =
    bytesToBase64Url(
      new TextEncoder().encode(
        JSON.stringify(payload)
      )
    );

  const key =
    await getAdminSessionKey(env);

  const signature =
    new Uint8Array(
      await crypto.subtle.sign(
        "HMAC",
        key,
        new TextEncoder().encode(
          encodedPayload
        )
      )
    );

  return {
    token:
      encodedPayload +
      "." +
      bytesToBase64Url(
        signature
      ),

    expiresAt:
      new Date(
        payload.exp * 1000
      ).toISOString(),
  };
}

async function verifyAdminSession(
  env,
  token
) {
  const parts =
    String(token || "")
      .split(".");

  if (parts.length !== 2) {
    throw httpError(
      401,
      "Session admin tidak valid"
    );
  }

  const [
    encodedPayload,
    encodedSignature
  ] = parts;

  let signature;

  try {
    signature =
      base64UrlToBytes(
        encodedSignature
      );
  } catch {
    throw httpError(
      401,
      "Session admin tidak valid"
    );
  }

  const key =
    await getAdminSessionKey(env);

  const valid =
    await crypto.subtle.verify(
      "HMAC",
      key,
      signature,
      new TextEncoder().encode(
        encodedPayload
      )
    );

  if (!valid) {
    throw httpError(
      401,
      "Session admin tidak valid"
    );
  }

  let payload;

  try {
    payload =
      JSON.parse(
        new TextDecoder().decode(
          base64UrlToBytes(
            encodedPayload
          )
        )
      );
  } catch {
    throw httpError(
      401,
      "Session admin tidak valid"
    );
  }

  const now =
    Math.floor(
      Date.now() / 1000
    );

  if (
    !payload.adminId ||
    !payload.exp ||
    payload.exp <= now
  ) {
    throw httpError(
      401,
      "Session admin sudah kedaluwarsa"
    );
  }

  return payload;
}

function getAdminBearerToken(
  request
) {
  const header =
    request.headers.get(
      "Authorization"
    ) || "";

  const match =
    header.match(
      /^Bearer\s+(.+)$/i
    );

  return match
    ? match[1].trim()
    : "";
}

async function findAdminById(
  env,
  adminId
) {
  const accessToken =
    await getGoogleAccessToken(env);

  const rows =
    await readValues(
      accessToken,
      getSheetId(env),
      "ADMINS!A2:H"
    );

  const index =
    rows.findIndex(
      (row) =>
        String(row[0] || "") ===
        String(adminId || "")
    );

  if (index === -1) {
    return null;
  }

  const row =
    rows[index];

  return {
    rowNumber:
      index + 2,

    adminId:
      String(row[0] || ""),

    username:
      String(row[1] || ""),

    displayName:
      String(row[2] || ""),

    passwordHash:
      String(row[3] || ""),

    passwordSalt:
      String(row[4] || ""),

    active:
      String(row[5] || "")
        .trim()
        .toUpperCase() ===
      "TRUE",

    createdAt:
      String(row[6] || ""),

    lastLoginAt:
      String(row[7] || ""),
  };
}

async function getAuthenticatedAdmin(
  env,
  request
) {
  const token =
    getAdminBearerToken(
      request
    );

  if (!token) {
    throw httpError(
      401,
      "Login admin diperlukan"
    );
  }

  const session =
    await verifyAdminSession(
      env,
      token
    );

  const admin =
    await findAdminById(
      env,
      session.adminId
    );

  if (!admin) {
    throw httpError(
      401,
      "Akun admin tidak ditemukan"
    );
  }

  if (!admin.active) {
    throw httpError(
      403,
      "Akun admin tidak aktif"
    );
  }

  return {
    adminId:
      admin.adminId,
    username:
      admin.username,
    displayName:
      admin.displayName,
  };
}

async function updateAdminLastLogin(
  env,
  rowNumber
) {
  const accessToken =
    await getGoogleAccessToken(env);

  const sheetId =
    getSheetId(env);

  const value =
    new Date().toISOString();

  const response =
    await fetch(
      "https://sheets.googleapis.com/v4/spreadsheets/" +
        sheetId +
        "/values/ADMINS!H" +
        rowNumber +
        "?valueInputOption=RAW",
      {
        method: "PUT",
        headers: {
          Authorization:
            "Bearer " +
            accessToken,

          "Content-Type":
            "application/json",
        },
        body: JSON.stringify({
          values: [
            [value],
          ],
        }),
      }
    );

  if (!response.ok) {
    const text =
      await response.text();

    throw new Error(
      "Gagal memperbarui LAST_LOGIN_AT: " +
      text
    );
  }
}

async function loginAdmin(
  env,
  username,
  password
) {
  username =
    cleanText(username)
      .toLowerCase();

  password =
    String(password || "");

  if (
    !username ||
    !password
  ) {
    throw httpError(
      400,
      "Username dan password wajib diisi"
    );
  }

  const accessToken =
    await getGoogleAccessToken(env);

  const rows =
    await readValues(
      accessToken,
      getSheetId(env),
      "ADMINS!A2:H"
    );

  const index =
    rows.findIndex(
      (row) =>
        String(row[1] || "")
          .trim()
          .toLowerCase() ===
        username
    );

  if (index === -1) {
    throw httpError(
      401,
      "Username atau password salah"
    );
  }

  const row =
    rows[index];

  const active =
    String(row[5] || "")
      .trim()
      .toUpperCase() ===
    "TRUE";

  if (!active) {
    throw httpError(
      403,
      "Akun admin tidak aktif"
    );
  }

  const passwordValid =
    await verifyAdminPassword(
      password,
      row[3],
      row[4]
    );

  if (!passwordValid) {
    throw httpError(
      401,
      "Username atau password salah"
    );
  }

  const admin = {
    adminId:
      String(row[0] || ""),

    username:
      String(row[1] || ""),

    displayName:
      String(row[2] || ""),
  };

  const session =
    await createAdminSession(
      env,
      admin
    );

  await updateAdminLastLogin(
    env,
    index + 2
  );

  return {
    admin,
    sessionToken:
      session.token,
    expiresAt:
      session.expiresAt,
  };
}



async function getAdminDashboard(
  env,
  request
) {
  const admin =
    await getAuthenticatedAdmin(
      env,
      request
    );

  const accessToken =
    await getGoogleAccessToken(env);

  const sheetId =
    getSheetId(env);

  const [
    submissionRows,
    companyRows,
    auditRows
  ] =
    await Promise.all([
      readValues(
        accessToken,
        sheetId,
        "SUBMISSIONS!A2:N"
      ),

      readValues(
        accessToken,
        sheetId,
        "COMPANIES!A2:E"
      ),

      readValues(
        accessToken,
        sheetId,
        "AUDIT_LOG!A2:G"
      ),
    ]);

  const companyMap =
    new Map(
      companyRows.map(
        (row) => [
          String(row[0] || ""),
          String(row[1] || ""),
        ]
      )
    );

  const submissionMap =
    new Map(
      submissionRows.map(
        (row) => [
          String(row[0] || ""),
          String(row[1] || ""),
        ]
      )
    );

  const stats = {
    total:
      submissionRows.length,

    menungguAsistensi: 0,
    perluPerbaikan: 0,
    final: 0,
    selesai: 0,
  };

  for (
    const row of submissionRows
  ) {
    const status =
      String(row[6] || "")
        .trim()
        .toUpperCase();

    if (
      status ===
      "MENUNGGU_ASISTENSI"
    ) {
      stats.menungguAsistensi++;
    }
    else if (
      status ===
      "PERLU_PERBAIKAN"
    ) {
      stats.perluPerbaikan++;
    }
    else if (
      status === "FINAL"
    ) {
      stats.final++;
    }
    else if (
      status === "SELESAI"
    ) {
      stats.selesai++;
    }
  }

  const recentSubmissions =
    submissionRows
      .map(
        (row) => ({
          submissionId:
            String(row[0] || ""),

          submissionNumber:
            String(row[1] || ""),

          companyName:
            companyMap.get(
              String(row[2] || "")
            ) || "-",

          picName:
            String(row[3] || ""),

          status:
            String(row[6] || ""),

          activeVersion:
            Number(row[9] || 0),

          createdAt:
            String(row[10] || ""),

          updatedAt:
            String(row[11] || ""),
        })
      )
      .sort(
        (a, b) =>
          new Date(
            b.updatedAt ||
            b.createdAt ||
            0
          ) -
          new Date(
            a.updatedAt ||
            a.createdAt ||
            0
          )
      )
      .slice(0, 5);

  const recentActivity =
    auditRows
      .map(
        (row) => ({
          logId:
            String(row[0] || ""),

          submissionId:
            String(row[1] || ""),

          submissionNumber:
            submissionMap.get(
              String(row[1] || "")
            ) || "-",

          action:
            String(row[2] || ""),

          actorType:
            String(row[3] || ""),

          actorName:
            String(row[4] || ""),

          detail:
            String(row[5] || ""),

          createdAt:
            String(row[6] || ""),
        })
      )
      .sort(
        (a, b) =>
          new Date(
            b.createdAt || 0
          ) -
          new Date(
            a.createdAt || 0
          )
      )
      .filter(
        (item) =>
          item.actorType
            .toUpperCase() ===
          "ADMIN"
      )
      .slice(0, 8);

  return {
    admin,
    stats,
    recentSubmissions,
    recentActivity,
  };
}



async function getAdminSubmissions(
  env,
  request,
  url
) {
  const admin =
    await getAuthenticatedAdmin(
      env,
      request
    );

  const accessToken =
    await getGoogleAccessToken(env);

  const sheetId =
    getSheetId(env);

  const [
    submissionRows,
    companyRows
  ] =
    await Promise.all([
      readValues(
        accessToken,
        sheetId,
        "SUBMISSIONS!A2:N"
      ),

      readValues(
        accessToken,
        sheetId,
        "COMPANIES!A2:E"
      ),
    ]);

  const companyMap =
    new Map(
      companyRows.map(
        (row) => [
          String(row[0] || ""),
          String(row[1] || ""),
        ]
      )
    );

  const query =
    String(
      url.searchParams.get("q") || ""
    )
      .trim()
      .toLowerCase();

  const statusFilter =
    String(
      url.searchParams.get("status") || ""
    )
      .trim()
      .toUpperCase();

  const validStatuses =
    new Set([
      "MENUNGGU_ASISTENSI",
      "PERLU_PERBAIKAN",
      "FINAL",
      "SELESAI",
    ]);

  if (
    statusFilter &&
    !validStatuses.has(
      statusFilter
    )
  ) {
    throw httpError(
      400,
      "Filter status tidak valid"
    );
  }

  let items =
    submissionRows.map(
      (row) => {
        const companyName =
          companyMap.get(
            String(row[2] || "")
          ) || "-";

        return {
          submissionId:
            String(row[0] || ""),

          submissionNumber:
            String(row[1] || ""),

          companyName,

          picName:
            String(row[3] || ""),

          picEmail:
            String(row[4] || ""),

          status:
            String(row[6] || ""),

          activeFileName:
            String(row[8] || ""),

          activeVersion:
            Number(row[9] || 0),

          createdAt:
            String(row[10] || ""),

          updatedAt:
            String(row[11] || ""),

          finalAt:
            String(row[12] || ""),

          completedAt:
            String(row[13] || ""),
        };
      }
    );

  if (statusFilter) {
    items =
      items.filter(
        (item) =>
          item.status ===
          statusFilter
      );
  }

  if (query) {
    items =
      items.filter(
        (item) => {
          const haystack = [
            item.submissionNumber,
            item.companyName,
            item.picName,
            item.picEmail,
          ]
            .join(" ")
            .toLowerCase();

          return haystack.includes(
            query
          );
        }
      );
  }

  items.sort(
    (a, b) =>
      new Date(
        b.updatedAt ||
        b.createdAt ||
        0
      ) -
      new Date(
        a.updatedAt ||
        a.createdAt ||
        0
      )
  );

  return {
    admin,
    total:
      items.length,
    items,
  };
}



async function getAdminSubmissionDetail(
  env,
  request,
  submissionNumber
) {
  const admin =
    await getAuthenticatedAdmin(
      env,
      request
    );

  submissionNumber =
    cleanText(
      submissionNumber
    );

  if (!submissionNumber) {
    throw httpError(
      400,
      "Nomor pengajuan wajib diisi"
    );
  }

  const accessToken =
    await getGoogleAccessToken(env);

  const sheetId =
    getSheetId(env);

  const [
    submissionRows,
    companyRows,
    versionRows,
    noteRows
  ] =
    await Promise.all([
      readValues(
        accessToken,
        sheetId,
        "SUBMISSIONS!A2:N"
      ),

      readValues(
        accessToken,
        sheetId,
        "COMPANIES!A2:E"
      ),

      readValues(
        accessToken,
        sheetId,
        "VERSIONS!A2:I"
      ),

      readValues(
        accessToken,
        sheetId,
        "ASSISTANCE_NOTES!A2:F"
      ),
    ]);

  const row =
    submissionRows.find(
      (item) =>
        String(item[1] || "")
          .trim()
          .toUpperCase() ===
        submissionNumber
          .trim()
          .toUpperCase()
    );

  if (!row) {
    throw httpError(
      404,
      "Pengajuan tidak ditemukan"
    );
  }

  const submissionId =
    String(row[0] || "");

  const companyId =
    String(row[2] || "");

  const companyRow =
    companyRows.find(
      (item) =>
        String(item[0] || "") ===
        companyId
    );

  const status =
    String(row[6] || "");

  const versions =
    versionRows
      .filter(
        (item) =>
          String(item[1] || "") ===
          submissionId
      )
      .map(
        (item) => ({
          versionId:
            String(item[0] || ""),

          versionNumber:
            Number(item[2] || 0),

          fileName:
            String(item[4] || ""),

          fileSizeBytes:
            Number(item[5] || 0),

          uploadedBy:
            String(item[6] || ""),

          uploadedAt:
            String(item[7] || ""),

          fileDeletedAt:
            String(item[8] || ""),

          isActive:
            Number(item[2] || 0) ===
            Number(row[9] || 0),
        })
      )
      .sort(
        (a, b) =>
          b.versionNumber -
          a.versionNumber
      );

  const notes =
    noteRows
      .filter(
        (item) =>
          String(item[1] || "") ===
          submissionId
      )
      .map(
        (item) => ({
          noteId:
            String(item[0] || ""),

          versionNumber:
            Number(item[2] || 0),

          noteText:
            String(item[3] || ""),

          createdBy:
            String(item[4] || ""),

          createdAt:
            String(item[5] || ""),
        })
      )
      .sort(
        (a, b) =>
          new Date(
            b.createdAt || 0
          ) -
          new Date(
            a.createdAt || 0
          )
      );

  return {
    admin,

    submission: {
      submissionId,

      submissionNumber:
        String(row[1] || ""),

      companyName:
        String(
          companyRow?.[1] ||
          "-"
        ),

      picName:
        String(row[3] || ""),

      picEmail:
        String(row[4] || ""),

      status,

      activeFileName:
        String(row[8] || ""),

      activeVersion:
        Number(row[9] || 0),

      fileAvailable:
        Boolean(
          String(row[7] || "")
        ) &&
        status !== "SELESAI",

      createdAt:
        String(row[10] || ""),

      updatedAt:
        String(row[11] || ""),

      finalAt:
        String(row[12] || ""),

      completedAt:
        String(row[13] || ""),
    },

    versions,
    notes,
  };
}

async function getAdminSubmissionFile(
  env,
  request,
  submissionNumber
) {
  await getAuthenticatedAdmin(
    env,
    request
  );

  submissionNumber =
    cleanText(
      submissionNumber
    );

  const accessToken =
    await getGoogleAccessToken(env);

  const rows =
    await readValues(
      accessToken,
      getSheetId(env),
      "SUBMISSIONS!A2:N"
    );

  const row =
    rows.find(
      (item) =>
        String(item[1] || "")
          .trim()
          .toUpperCase() ===
        submissionNumber
          .trim()
          .toUpperCase()
    );

  if (!row) {
    throw httpError(
      404,
      "Pengajuan tidak ditemukan"
    );
  }

  const status =
    String(row[6] || "");

  const fileId =
    String(row[7] || "");

  const fileName =
    String(row[8] || "") ||
    submissionNumber + ".pdf";

  if (
    status === "SELESAI" ||
    !fileId
  ) {
    throw httpError(
      410,
      "Berkas aktif sudah tidak tersedia"
    );
  }

  const driveResponse =
    await fetch(
      "https://www.googleapis.com/drive/v3/files/" +
        encodeURIComponent(fileId) +
        "?alt=media",
      {
        headers: {
          Authorization:
            "Bearer " +
            accessToken,
        },
      }
    );

  if (!driveResponse.ok) {
    const text =
      await driveResponse.text();

    throw new Error(
      "Gagal mengambil PDF dari Google Drive: " +
      text
    );
  }

  const buffer =
    await driveResponse.arrayBuffer();

  return {
    buffer,
    fileName,
  };
}


export default {
  async fetch(request, env) {
    const url =
      new URL(request.url);

    const corsHeaders =
      getCorsHeaders(request);

    if (
      request.method === "OPTIONS"
    ) {
      return new Response(
        null,
        {
          status: 204,
          headers:
            corsHeaders,
        }
      );
    }

    if (
      request.method === "GET" &&
      url.pathname === "/health"
    ) {
      return jsonResponse(
        request,
        {
          ok: true,
          app: "ASBUN",
          name:
            "Asistensi Berkas Usaha dan Perizinan",
          version:
            "1.0.1",
        }
      );
    }

    if (
      request.method === "POST" &&
      url.pathname ===
        "/submissions"
    ) {
      try {
        const parsed =
          await parseSubmission(
            request
          );

        const result =
          await createSubmission(
            env,
            parsed
          );

        return jsonResponse(
          request,
          {
            ok: true,
            message:
              "Pengajuan berhasil diterima",
            submission:
              result,
          },
          201
        );
      } catch (error) {
        return jsonResponse(
          request,
          {
            ok: false,
            error:
              error.message ||
              "Gagal membuat pengajuan",
          },
          500
        );
      }
    }

    const revisionMatch =
      url.pathname.match(
        /^\/admin\/submissions\/([^/]+)\/needs-revision$/
      );

    if (
      request.method === "POST" &&
      revisionMatch
    ) {
      try {
        const admin =
          await getAuthenticatedAdmin(
            env,
            request
          );

        let body;

        try {
          body = await request.json();
        } catch {
          throw httpError(
            400,
            "Body harus berupa JSON"
          );
        }

        const submissionNumber =
          decodeURIComponent(
            revisionMatch[1]
          );

        const result =
          await markNeedsRevision(
            env,
            {
              submissionNumber,
              note: body.note,
              admin,
            }
          );

        return jsonResponse(
          request,
          {
            ok: true,
            message:
              "Pengajuan ditandai PERLU_PERBAIKAN",
            submission:
              result,
          }
        );
      } catch (error) {
        return jsonResponse(
          request,
          {
            ok: false,
            error:
              error.message ||
              "Gagal memproses asistensi",
          },
          error.status || 500
        );
      }
    }


    if (
      request.method === "POST" &&
      url.pathname ===
        "/private/revision"
    ) {
      try {
        const parsed =
          await parseRevisionUpload(
            request
          );

        const result =
          await uploadRevision(
            env,
            parsed
          );

        return jsonResponse(
          request,
          {
            ok: true,
            message:
              "Versi perbaikan berhasil diunggah",
            submission:
              result,
          }
        );
      } catch (error) {
        return jsonResponse(
          request,
          {
            ok: false,
            error:
              error.message ||
              "Gagal mengunggah versi perbaikan",
          },
          error.status || 500
        );
      }
    }


    const finalMatch =
      url.pathname.match(
        /^\/admin\/submissions\/([^/]+)\/final$/
      );

    if (
      request.method === "POST" &&
      finalMatch
    ) {
      try {
        const admin =
          await getAuthenticatedAdmin(
            env,
            request
          );

        const result =
          await markSubmissionFinal(
            env,
            {
              submissionNumber:
                decodeURIComponent(
                  finalMatch[1]
                ),
              admin,
            }
          );

        return jsonResponse(
          request,
          {
            ok: true,
            message:
              "Pengajuan ditandai FINAL",
            submission:
              result,
          }
        );
      } catch (error) {
        return jsonResponse(
          request,
          {
            ok: false,
            error:
              error.message ||
              "Gagal menetapkan FINAL",
          },
          error.status || 500
        );
      }
    }

    const completeMatch =
      url.pathname.match(
        /^\/admin\/submissions\/([^/]+)\/complete$/
      );

    if (
      request.method === "POST" &&
      completeMatch
    ) {
      try {
        const admin =
          await getAuthenticatedAdmin(
            env,
            request
          );

        const result =
          await completeSubmission(
            env,
            {
              submissionNumber:
                decodeURIComponent(
                  completeMatch[1]
                ),
              admin,
            }
          );

        return jsonResponse(
          request,
          {
            ok: true,
            message:
              "Pengajuan selesai dan berkas telah dihapus",
            submission:
              result,
          }
        );
      } catch (error) {
        return jsonResponse(
          request,
          {
            ok: false,
            error:
              error.message ||
              "Gagal menyelesaikan pengajuan",
          },
          error.status || 500
        );
      }
    }


    if (
      request.method === "GET" &&
      url.pathname ===
        "/private/submission"
    ) {
      try {
        const privateToken =
          url.searchParams.get("t");

        const result =
          await getPrivateSubmissionView(
            env,
            privateToken
          );

        return jsonResponse(
          request,
          {
            ok: true,
            submission:
              result,
          }
        );
      } catch (error) {
        return jsonResponse(
          request,
          {
            ok: false,
            error:
              error.message ||
              "Gagal membaca pengajuan",
          },
          error.status || 500
        );
      }
    }


    if (
      request.method === "GET" &&
      url.pathname ===
        "/private/file"
    ) {
      try {
        const result =
          await getPrivateActiveFile(
            env,
            url.searchParams.get("t")
          );

        const safeFileName =
          String(result.fileName)
            .replace(/[\r\n"]/g, "_");

        return new Response(
          result.buffer,
          {
            status: 200,
            headers: {
              ...getCorsHeaders(request),
              "Content-Type":
                "application/pdf",
              "Content-Disposition":
                "inline; filename=\"" +
                safeFileName +
                "\"",
              "Cache-Control":
                "private, no-store",
              "X-Content-Type-Options":
                "nosniff",
              "Access-Control-Expose-Headers":
                "Content-Disposition",
            },
          }
        );
      } catch (error) {
        return jsonResponse(
          request,
          {
            ok: false,
            error:
              error.message ||
              "Gagal membuka berkas",
          },
          error.status || 500
        );
      }
    }


    if (
      request.method === "POST" &&
      url.pathname === "/admin/login"
    ) {
      try {
        let body;

        try {
          body =
            await request.json();
        } catch {
          throw httpError(
            400,
            "Body harus berupa JSON"
          );
        }

        const result =
          await loginAdmin(
            env,
            body.username,
            body.password
          );

        return jsonResponse(
          request,
          {
            ok: true,
            admin:
              result.admin,
            sessionToken:
              result.sessionToken,
            expiresAt:
              result.expiresAt,
          }
        );
      } catch (error) {
        return jsonResponse(
          request,
          {
            ok: false,
            error:
              error.message ||
              "Login admin gagal",
          },
          error.status || 500
        );
      }
    }

    if (
      request.method === "GET" &&
      url.pathname === "/admin/me"
    ) {
      try {
        const admin =
          await getAuthenticatedAdmin(
            env,
            request
          );

        return jsonResponse(
          request,
          {
            ok: true,
            admin,
          }
        );
      } catch (error) {
        return jsonResponse(
          request,
          {
            ok: false,
            error:
              error.message ||
              "Session admin tidak valid",
          },
          error.status || 500
        );
      }
    }

    if (
      request.method === "GET" &&
      url.pathname === "/admin/dashboard"
    ) {
      try {
        const dashboard =
          await getAdminDashboard(
            env,
            request
          );

        return jsonResponse(
          request,
          {
            ok: true,
            ...dashboard,
          }
        );
      } catch (error) {
        return jsonResponse(
          request,
          {
            ok: false,
            error:
              error.message ||
              "Gagal memuat dashboard admin",
          },
          error.status || 500
        );
      }
    }

    if (
      request.method === "GET" &&
      url.pathname === "/admin/submissions"
    ) {
      try {
        const result =
          await getAdminSubmissions(
            env,
            request,
            url
          );

        return jsonResponse(
          request,
          {
            ok: true,
            ...result,
          }
        );
      } catch (error) {
        return jsonResponse(
          request,
          {
            ok: false,
            error:
              error.message ||
              "Gagal memuat antrean pengajuan",
          },
          error.status || 500
        );
      }
    }

    const adminFileMatch =
      url.pathname.match(
        /^\/admin\/submissions\/([^/]+)\/file$/
      );

    if (
      request.method === "GET" &&
      adminFileMatch
    ) {
      try {
        const submissionNumber =
          decodeURIComponent(
            adminFileMatch[1]
          );

        const file =
          await getAdminSubmissionFile(
            env,
            request,
            submissionNumber
          );

        return new Response(
          file.buffer,
          {
            status: 200,
            headers: {
              ...corsHeaders,

              "Content-Type":
                "application/pdf",

              "Content-Disposition":
                'inline; filename="' +
                file.fileName
                  .replace(/"/g, "") +
                '"',

              "Cache-Control":
                "private, no-store",

              "X-Content-Type-Options":
                "nosniff",
            },
          }
        );
      } catch (error) {
        return jsonResponse(
          request,
          {
            ok: false,
            error:
              error.message ||
              "Gagal membuka PDF",
          },
          error.status || 500
        );
      }
    }

    const adminDetailMatch =
      url.pathname.match(
        /^\/admin\/submissions\/([^/]+)$/
      );

    if (
      request.method === "GET" &&
      adminDetailMatch
    ) {
      try {
        const submissionNumber =
          decodeURIComponent(
            adminDetailMatch[1]
          );

        const result =
          await getAdminSubmissionDetail(
            env,
            request,
            submissionNumber
          );

        return jsonResponse(
          request,
          {
            ok: true,
            ...result,
          }
        );
      } catch (error) {
        return jsonResponse(
          request,
          {
            ok: false,
            error:
              error.message ||
              "Gagal memuat detail pengajuan",
          },
          error.status || 500
        );
      }
    }

    return jsonResponse(
      request,
      {
        ok: false,
        error: "NOT_FOUND",
      },
      404
    );
  },
};
