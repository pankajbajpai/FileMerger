const EXT_BY_TYPE = {
  pdf: ["pdf"],
  excel: ["xlsx", "xls", "csv"],
  word: ["docx"],
};

let currentType = "pdf";
let selectedFiles = []; // { file, protected, password }

const typeBtns = document.querySelectorAll(".type-btn");
const fileInput = document.getElementById("fileInput");
const fileListEl = document.getElementById("fileList");
const outNameEl = document.getElementById("outName");
const mergeBtn = document.getElementById("mergeBtn");
const statusEl = document.getElementById("status");

function acceptAttr() {
  return EXT_BY_TYPE[currentType].map((e) => "." + e).join(",");
}

typeBtns.forEach((btn) => {
  btn.addEventListener("click", () => {
    typeBtns.forEach((b) => b.classList.remove("active"));
    btn.classList.add("active");
    currentType = btn.dataset.type;
    fileInput.setAttribute("accept", acceptAttr());
    selectedFiles = [];
    renderFileList();
    setStatus("", "");
  });
});
fileInput.setAttribute("accept", acceptAttr());

fileInput.addEventListener("change", () => {
  const allowed = EXT_BY_TYPE[currentType];
  const rejected = [];
  for (const file of fileInput.files) {
    const ext = file.name.split(".").pop().toLowerCase();
    if (!allowed.includes(ext)) {
      rejected.push(file.name);
      continue;
    }
    selectedFiles.push({ file, protected: false, password: "" });
  }
  fileInput.value = "";
  if (rejected.length) {
    setStatus(
      `Skipped file(s) not matching "${currentType}" type: ${rejected.join(", ")}`,
      "error"
    );
  }
  renderFileList();
});

function renderFileList() {
  fileListEl.innerHTML = "";
  selectedFiles.forEach((entry, idx) => {
    const row = document.createElement("div");
    row.className = "file-row" + (entry.protected ? " protected" : "");

    const name = document.createElement("span");
    name.className = "fname";
    name.textContent = `${idx + 1}. ${entry.file.name}`;
    row.appendChild(name);

    const label = document.createElement("label");
    label.className = "protected-label";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = entry.protected;
    cb.addEventListener("change", () => {
      entry.protected = cb.checked;
      renderFileList();
    });
    label.appendChild(cb);
    label.appendChild(document.createTextNode("Password protected"));
    row.appendChild(label);

    const pw = document.createElement("input");
    pw.type = "password";
    pw.className = "pw-input";
    pw.placeholder = "Password";
    pw.value = entry.password;
    pw.addEventListener("input", () => {
      entry.password = pw.value;
    });
    row.appendChild(pw);

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

mergeBtn.addEventListener("click", async () => {
  if (selectedFiles.length < 2) {
    setStatus("Add at least two files of the selected type to merge.", "error");
    return;
  }

  const passwords = {};
  selectedFiles.forEach((entry, idx) => {
    if (entry.protected && entry.password) {
      passwords[idx] = entry.password;
    }
  });

  const missing = selectedFiles.some((e) => e.protected && !e.password);
  if (missing) {
    setStatus("Please enter a password for every file marked as protected.", "error");
    return;
  }

  const formData = new FormData();
  formData.append("file_type", currentType);
  formData.append("output_name", outNameEl.value.trim());
  formData.append("passwords", JSON.stringify(passwords));
  selectedFiles.forEach((entry) => formData.append("files", entry.file));

  mergeBtn.disabled = true;
  setStatus("Merging…", "");

  try {
    const resp = await fetch("/api/merge", { method: "POST", body: formData });
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

    setStatus(`Done — "${filename}" downloaded. Choose the save location in your browser's download prompt if asked.`, "ok");
  } catch (err) {
    setStatus("Network error: " + err.message, "error");
  } finally {
    mergeBtn.disabled = false;
  }
});
