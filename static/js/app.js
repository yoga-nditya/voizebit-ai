const textInput = document.getElementById("textInput");
const sendBtn = document.getElementById("sendBtn");
const sendCreateBtn = document.getElementById("sendCreateBtn");
const chatBox = document.getElementById("chatBox");
const filenameInput = document.getElementById("filename");
const optDocx = document.getElementById("optDocx");
const optXlsx = document.getElementById("optXlsx");
const optPptx = document.getElementById("optPptx");
const filesLinks = document.getElementById("filesLinks");

async function sendMessage(text, options = { auto_create: false, filename: null, types: ["docx"] }) {
  chatBox.innerText = "Thinking...";
  filesLinks.innerHTML = "";
  try {
    const body = {
      message: text,
      auto_create: options.auto_create
    };
    if (options.auto_create && options.filename) body.filename = options.filename;
    if (options.auto_create && options.types) body.types = options.types;

    const res = await fetch("/api/chat", {
      method: "POST",
      headers: {"Content-Type":"application/json"},
      body: JSON.stringify(body)
    });
    const data = await res.json();
    if (data.error) {
      chatBox.innerText = "Error: " + data.error + (data.detail ? ("\n"+data.detail) : "");
      return;
    }
    chatBox.innerText = data.text;
    if (data.files && data.files.length) {
      filesLinks.innerHTML = "";
      data.files.forEach(f => {
        const a = document.createElement("a");
        a.href = f.url;
        a.textContent = `Download ${f.filename}`;
        a.target = "_blank";
        filesLinks.appendChild(a);
      });
    }
  } catch (e) {
    chatBox.innerText = "Fetch error: " + e.toString();
  }
}

sendBtn.addEventListener("click", () => {
  const t = textInput.value.trim();
  if (!t) return alert("Ketik prompt dulu.");
  sendMessage(t, { auto_create: false });
});

sendCreateBtn.addEventListener("click", () => {
  const t = textInput.value.trim();
  if (!t) return alert("Ketik prompt dulu.");
  const fname = (filenameInput.value || "document").trim();
  const types = [];
  if (optDocx.checked) types.push("docx");
  if (optXlsx.checked) types.push("xlsx");
  if (optPptx.checked) types.push("pptx");
  if (types.length === 0) return alert("Pilih minimal satu tipe file (DOCX/XLSX/PPTX).");
  sendMessage(t, { auto_create: true, filename: fname, types: types });
});
