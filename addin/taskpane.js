const API_URL = "https://automatizacion-dr22.onrender.com/procesar-plan-base64";

let finalHtmlTables = "";

Office.onReady(() => {
  document.getElementById("processBtn").addEventListener("click", processPdfAttachments);
  document.getElementById("copyBtn").addEventListener("click", copyHtmlTables);
  setStatus("Listo.");
});

function setStatus(message, type = "") {
  const el = document.getElementById("status");
  el.textContent = message;
  el.className = "status";

  if (type === "ok") el.classList.add("ok");
  if (type === "error") el.classList.add("error");
}

function setPreview(html) {
  document.getElementById("preview").innerHTML = html;
}

function getCurrentItem() {
  return Office.context.mailbox.item;
}

function getAttachments(item) {
  return new Promise((resolve, reject) => {
    if (typeof item.getAttachmentsAsync === "function") {
      item.getAttachmentsAsync(result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || []);
        } else {
          reject(new Error(result.error.message || "No se pudieron leer los adjuntos."));
        }
      });
      return;
    }

    resolve(item.attachments || []);
  });
}

async function getPdfAttachments(item) {
  const attachments = await getAttachments(item);

  return attachments.filter(att => {
    const name = (att.name || "").toLowerCase();
    return name.endsWith(".pdf");
  });
}

function getAttachmentContent(item, attachmentId) {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(attachmentId, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message || "No se pudo leer el adjunto."));
      }
    });
  });
}

async function sendPdfToApi(filename, base64Content) {
  const response = await fetch(API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      filename: filename,
      content_base64: base64Content
    })
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`API error ${response.status}: ${errorText}`);
  }

  return await response.json();
}

async function processPdfAttachments() {
  const processBtn = document.getElementById("processBtn");
  const copyBtn = document.getElementById("copyBtn");

  processBtn.disabled = true;
  copyBtn.disabled = true;
  finalHtmlTables = "";
  setPreview("");

  try {
    setStatus("Buscando adjuntos PDF...");

    const item = getCurrentItem();
    const pdfAttachments = await getPdfAttachments(item);

    if (pdfAttachments.length === 0) {
      setStatus("No encontré PDFs adjuntos en este correo.", "error");
      return;
    }

    setStatus(`Encontré ${pdfAttachments.length} PDF(s). Procesando...`);

    const tables = [];
    const errors = [];

    for (let i = 0; i < pdfAttachments.length; i++) {
      const attachment = pdfAttachments[i];

      try {
        setStatus(`Procesando ${i + 1}/${pdfAttachments.length}: ${attachment.name}`);

        const content = await getAttachmentContent(item, attachment.id);

        if (!content || !content.content) {
          throw new Error("El adjunto no devolvió contenido.");
        }

        /*
          Normalmente getAttachmentContentAsync devuelve:
          {
            content: "...base64...",
            format: "base64"
          }

          Si fuera un adjunto en la nube o formato no compatible, puede fallar.
        */
        const result = await sendPdfToApi(attachment.name, content.content);

        if (result.html_table) {
          tables.push(result.html_table);
        } else {
          errors.push(`${attachment.name}: la API respondió sin html_table.`);
        }

        if (result.advertencias && result.advertencias.length > 0) {
          errors.push(`${attachment.name}: ${result.advertencias.join(" | ")}`);
        }

      } catch (err) {
        errors.push(`${attachment.name}: ${err.message}`);
      }
    }

    if (tables.length === 0) {
      setStatus("No se pudo generar ninguna tabla.\n\n" + errors.join("\n"), "error");
      return;
    }

    finalHtmlTables = tables.join("<br><br>");
    setPreview(finalHtmlTables);

    copyBtn.disabled = false;

    let message = `Proceso terminado. Tablas generadas: ${tables.length}.`;
    if (errors.length > 0) {
      message += `\n\nAdvertencias:\n${errors.join("\n")}`;
    }

    setStatus(message, errors.length > 0 ? "error" : "ok");

  } catch (err) {
    setStatus("Error general: " + err.message, "error");
  } finally {
    processBtn.disabled = false;
  }
}

async function copyHtmlTables() {
  if (!finalHtmlTables) {
    setStatus("No hay tablas para copiar.", "error");
    return;
  }

  const emailHtml = `
    Saludos<br>
    Detallo información de:<br><br>
    ${finalHtmlTables}
  `;

  try {
    if (navigator.clipboard && window.ClipboardItem) {
      const blobHtml = new Blob([emailHtml], { type: "text/html" });
      const blobText = new Blob([stripHtml(emailHtml)], { type: "text/plain" });

      await navigator.clipboard.write([
        new ClipboardItem({
          "text/html": blobHtml,
          "text/plain": blobText
        })
      ]);

      setStatus("Tablas copiadas. Pégalas en el correo.", "ok");
    } else {
      await navigator.clipboard.writeText(stripHtml(emailHtml));
      setStatus("Copiado como texto. Si no conserva formato, copia manualmente desde la vista previa.", "ok");
    }
  } catch (err) {
    setStatus("No pude copiar automáticamente. Selecciona la tabla en la vista previa y cópiala manualmente.", "error");
  }
}

function stripHtml(html) {
  const div = document.createElement("div");
  div.innerHTML = html;
  return div.textContent || div.innerText || "";
}
