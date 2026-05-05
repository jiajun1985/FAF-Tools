(() => {
  function setupUpload({ input, dropzone, nameNode, clearButton, parseWorkbook, onParsed, onCleared, onChange }) {
    input.addEventListener("change", async () => {
      if (input.files && input.files[0]) {
        await handleFile({ file: input.files[0], nameNode, clearButton, parseWorkbook, onParsed, onChange });
      }
    });

    ["dragenter", "dragover"].forEach((eventName) => {
      dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        dropzone.classList.add("dragging");
      });
    });

    ["dragleave", "drop"].forEach((eventName) => {
      dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        dropzone.classList.remove("dragging");
      });
    });

    dropzone.addEventListener("drop", async (event) => {
      const file = event.dataTransfer.files && event.dataTransfer.files[0];
      if (file) {
        input.files = event.dataTransfer.files;
        await handleFile({ file, nameNode, clearButton, parseWorkbook, onParsed, onChange });
      }
    });

    clearButton.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      input.value = "";
      setFileState(nameNode, clearButton, "");
      onCleared();
      onChange();
    });
  }

  async function handleFile({ file, nameNode, clearButton, parseWorkbook, onParsed, onChange }) {
    setFileState(nameNode, clearButton, file.name);
    try {
      const parsed = await parseWorkbook(file);
      onParsed(parsed);
      onChange();
    } catch (error) {
      onParsed(null);
      onChange(error.message);
    }
  }

  function setFileState(nameNode, clearButton, fileName) {
    if (fileName) {
      nameNode.textContent = `已选择：${fileName}`;
      nameNode.classList.add("has-file");
      clearButton.classList.add("visible");
      return;
    }
    nameNode.textContent = "未选择文件";
    nameNode.classList.remove("has-file");
    clearButton.classList.remove("visible");
  }

  window.FAFULIFileUpload = {
    setupUpload,
    setFileState
  };
})();
