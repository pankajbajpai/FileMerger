const ALLOWED_EXTS = ["pdf", "xlsx", "xls", "csv", "docx"];

let selectedFiles = []; // { file, password, passwordRequired }

const fileInput = document.getElementById("fileInput");
const fileListEl = document.getElementById("fileList");
const outNameEl = document.getElementById("outName");
const mergeBtn = document.getElementById("mergeBtn");
const statusEl = document.getElementById("status");

fileInput.addEventListener("change", () => {
  const rejected = [];
  for (const file of fileInput.files) {
    const ext = file.name.split(".").pop().toLowerCase();
    if (!ALLOWED_EXTS.includes(ext)) {
      rejected.push(file.name);
      continue;
    }
    selectedFiles.push({ file, password: "", passwordRequired: false });
  }
  fileInput.value = "";
  if (rejected.length) {
    setStatus(
      `Skipped unsupported file(s): ${rejected.join(", ")}. Supported: ${ALLOWED_EXTS.map((e) => "." + e).join(", ")}`,
      "error"
    );
  } else {
    setStatus("", "");
  }
  renderFileList();
});

function renderFileList() {
  fileListEl.innerHTML = "";
  selectedFiles.forEach((entry, idx) => {
    const row = document.createElement("div");
    row.className = "file-row" + (entry.passwordRequired ? " protected" : "");

    const name = document.createElement("span");
    name.className = "fname";
    name.textContent = `${idx + 1}. ${entry.file.name}`;
    row.appendChild(name);

    if (entry.passwordRequired) {
      const pw = document.createElement("input");
      pw.type = "password";
      pw.className = "pw-input";
      pw.placeholder = "Password required";
      pw.value = entry.password;
      pw.addEventListener("input", () => {
        entry.password = pw.value;
      });
      row.appendChild(pw);
    }

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "remove-btn";
    removeBtn.textContent = "✕";
    removeBtn.addEventListener("click", () => {
      selectedFiles.splice(idx, 1);
      renderFileList();
    });
    row.appendChild(removeBtn);

    fileListEl.appendChild(row);
  });
}

function setStatus(msg, kind) {
  statusEl.textContent = msg;
  statusEl.className = "status" + (kind ? " " + kind : "");
}

async function attemptMerge() {
  if (selectedFiles.length < 1) {
    setStatus("Add at least one file.", "error");
    return;
  }

  const outName = outNameEl.value.trim();
  if (!outName || !outName.includes(".")) {
    setStatus('Output file name must include an extension, e.g. "merged.pdf".', "error");
    return;
  }

  const stillMissing = selectedFiles.filter((e) => e.passwordRequired && !e.password);
  if (stillMissing.length) {
    setStatus("Enter the password for the file(s) marked above.", "error");
    return;
  }

  const passwords = {};
  selectedFiles.forEach((entry, idx) => {
    if (entry.password) passwords[idx] = entry.password;
  });

  const formData = new FormData();
  formData.append("output_name", outName);
  formData.append("passwords", JSON.stringify(passwords));
  selectedFiles.forEach((entry) => formData.append("files", entry.file));

  mergeBtn.disabled = true;
  setStatus("Merging…", "");

  try {
    const resp = await fetch("/api/merge", { method: "POST", body: formData });

    if (resp.status === 422) {
      const data = await resp.json().catch(() => ({}));
      if (data.password_required && data.password_required.length) {
        data.password_required.forEach((item) => {
          selectedFiles[item.index].passwordRequired = true;
        });
        renderFileList();
        setStatus(
          `Password needed for: ${data.password_required.map((i) => i.name).join(", ")}. Enter it above and click Merge again.`,
          "error"
        );
        return;
      }
      setStatus(data.error || "Merge failed.", "error");
      return;
    }

    if (!resp.ok) {
      const data = await resp.json().catch(() => ({}));
      setStatus(data.error || `Merge failed (${resp.status}).`, "error");
      return;
    }

    const blob = await resp.blob();
    let filename = "merged_output";
    const disposition = resp.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="?([^"]+)"?/);
    if (match) filename = match[1];

    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);

    setStatus(
      `Done — "${filename}" downloaded. Use your browser's Save dialog to choose the destination if prompted.`,
      "ok"
    );
  } catch (err) {
    setStatus("Network error: " + err.message, "error");
  } finally {
    mergeBtn.disabled = false;
  }
}

mergeBtn.addEventListener("click", attemptMerge);
