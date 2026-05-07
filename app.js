const form = document.querySelector("#uploadForm");
const input = document.querySelector("#excelInput");
const dropZone = document.querySelector("#dropZone");
const fileName = document.querySelector("#fileName");
const statusText = document.querySelector("#status");
const submitButton = document.querySelector("#submitButton");

function setStatus(message, type = "") {
  statusText.textContent = message;
  statusText.className = `status ${type}`.trim();
}

function setFile(file) {
  if (!file) {
    fileName.textContent = "Choose or drop an .xlsx file here";
    return;
  }

  fileName.textContent = file.name;
  setStatus("");
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function outputName(inputName) {
  const base = inputName.replace(/\.[^.]+$/, "").replace(/[^a-z0-9_-]+/gi, "_") || "media_plan";
  return `${base}_media_plan.pptx`;
}

input.addEventListener("change", () => setFile(input.files[0]));

["dragenter", "dragover"].forEach(eventName => {
  dropZone.addEventListener(eventName, event => {
    event.preventDefault();
    dropZone.classList.add("is-dragging");
  });
});

["dragleave", "drop"].forEach(eventName => {
  dropZone.addEventListener(eventName, event => {
    event.preventDefault();
    dropZone.classList.remove("is-dragging");
  });
});

dropZone.addEventListener("drop", event => {
  const file = event.dataTransfer.files[0];
  if (!file) return;

  const transfer = new DataTransfer();
  transfer.items.add(file);
  input.files = transfer.files;
  setFile(file);
});

form.addEventListener("submit", async event => {
  event.preventDefault();

  const file = input.files[0];
  if (!file) {
    setStatus("Please choose an Excel file first.", "error");
    return;
  }

  if (!/\.(xlsx|xls)$/i.test(file.name)) {
    setStatus("Please upload an .xlsx or .xls file.", "error");
    return;
  }

  submitButton.disabled = true;
  submitButton.textContent = "Generating...";
  setStatus("Creating your PowerPoint in this browser. This may take a few seconds.");

  try {
    if (!window.MediaPlanGenerator) {
      throw new Error("The PPT generator did not load. Please refresh the page and try again.");
    }

    const excelBuffer = await file.arrayBuffer();
    const pptBlob = await window.MediaPlanGenerator.generateMediaPlanFromBuffer(excelBuffer);
    downloadBlob(pptBlob, outputName(file.name));
    setStatus("PPT generated and downloaded.", "success");
  } catch (error) {
    setStatus(error.message || "Could not generate the PPT.", "error");
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "Generate PPT";
  }
});
