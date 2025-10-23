document.addEventListener('DOMContentLoaded', () => {
  // 요소
  const uploadForm = document.getElementById('upload-form');
  const pdfFileInput = document.getElementById('pdf-file');
  const dropzone = document.getElementById('dropzone');
  const fileLabelText = document.getElementById('file-text');
  const fileNameDisplay = document.getElementById('file-name');
  const generateBtn = document.getElementById('generate-btn');
  const loadingDiv = document.getElementById('loading');
  const resultContainer = document.getElementById('result-container');
  const resultArea = document.getElementById('result-area');
  const rawArea = document.getElementById('raw-area');
  const errorContainer = document.getElementById('error-container');
  const errorMessage = document.getElementById('error-message');
  const copyBtn = document.getElementById('copy-btn');
  const downloadBtns = document.querySelectorAll('.download-btn');

  // ✅ API 엔드포인트
  const API_BASE = window.API_BASE || 'https://port-0-interviewpro-mh3iopw4cf627816.sel3.cloudtype.app';
  const GENERATE_URL = `${API_BASE}/generate`;
  const DOWNLOAD_URL = `${API_BASE}/download`;

  // 상태
  let selectedFile = null;
  let rawText = '';

  // 마크다운 옵션
  marked.use({ breaks: true, gfm: true });

  // ========== 드래그&드롭 ==========
  ['dragenter','dragover'].forEach(ev => {
    dropzone.addEventListener(ev, (e) => {
      e.preventDefault(); e.stopPropagation();
      dropzone.classList.add('dragover');
      fileLabelText.textContent = '여기에 놓으면 업로드됩니다!';
    });
  });
  ['dragleave','drop'].forEach(ev => {
    dropzone.addEventListener(ev, (e) => {
      e.preventDefault(); e.stopPropagation();
      dropzone.classList.remove('dragover');
      fileLabelText.textContent = '클릭하거나 여기로 PDF를 끌어다 놓으세요';
    });
  });
  dropzone.addEventListener('drop', (e) => {
    const items = e.dataTransfer.files;
    if (!items || !items.length) return;
    const file = items[0];
    if (!file || (file.type !== 'application/pdf' && !file.name.toLowerCase().endsWith('.pdf'))) {
      alert('PDF 파일만 업로드할 수 있습니다.');
      return;
    }
    selectedFile = file;
    fileNameDisplay.textContent = `파일명: ${file.name}`;
    generateBtn.disabled = false;
  });

  // input 선택
  pdfFileInput.addEventListener('change', () => {
    const file = pdfFileInput.files[0];
    if (file) {
      if (!file.type.startsWith('application/pdf') && !file.name.toLowerCase().endsWith('.pdf')) {
        alert('PDF 파일만 선택해주세요.');
        pdfFileInput.value = ''; selectedFile = null; generateBtn.disabled = true; return;
      }
      selectedFile = file;
      fileLabelText.textContent = '파일이 선택되었습니다!';
      fileNameDisplay.textContent = `파일명: ${file.name}`;
      generateBtn.disabled = false;
    } else {
      selectedFile = null;
      fileLabelText.textContent = '클릭하거나 여기로 PDF를 끌어다 놓으세요';
      fileNameDisplay.textContent = '';
      generateBtn.disabled = true;
    }
  });

  function hideResults() {
    resultContainer.classList.add('hidden');
    errorContainer.classList.add('hidden');
    resultArea.innerHTML = '';
    rawArea.textContent = '';
    rawText = '';
    setActionButtonsEnabled(false);
  }
  function setActionButtonsEnabled(enabled) {
    copyBtn.disabled = !enabled;
    downloadBtns.forEach(b => (b.disabled = !enabled));
  }
  function setLoadingState(isLoading) {
    if (isLoading) {
      generateBtn.disabled = true;
      generateBtn.textContent = '작성 중...';
      loadingDiv.classList.remove('hidden');
    } else {
      generateBtn.disabled = false;
      generateBtn.textContent = '✨ 질문/답변 생성 시작';
      loadingDiv.classList.add('hidden');
    }
  }
  function displayError(message) {
    errorMessage.textContent = message;
    errorContainer.classList.remove('hidden');
  }
  function renderMarkdownToResult(text) {
    const html = DOMPurify.sanitize(marked.parse(text || ''));
    resultArea.innerHTML = html;
  }
  function autoScrollResult() {
    resultArea.scrollTop = resultArea.scrollHeight;
  }

  // 제출(생성)
  uploadForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    hideResults();
    setLoadingState(true);

    if (!selectedFile) {
      setLoadingState(false);
      displayError('PDF 파일이 선택되지 않았습니다.');
      return;
    }

    try {
      const formData = new FormData();
      formData.append('file', selectedFile);

      const response = await fetch(GENERATE_URL, { method: 'POST', body: formData });
      if (!response.ok) throw new Error(`서버 오류: ${response.status} ${response.statusText}`);

      resultContainer.classList.remove('hidden');
      resultArea.innerHTML = '';
      rawText = '';
      rawArea.textContent = '';

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        const chunk = decoder.decode(value, { stream: true });
        rawText += chunk;
        rawArea.textContent = rawText;
        renderMarkdownToResult(rawText);
        autoScrollResult();
      }
      setActionButtonsEnabled(rawText.trim().length > 0);
    } catch (err) {
      console.error('[GENERATE fetch error]', err);
      displayError(err.message || '요청에 실패했습니다.');
    } finally {
      setLoadingState(false);
    }
  });

  // 복사
  copyBtn.addEventListener('click', async () => {
    try {
      if (!rawText.trim()) { alert('복사할 결과가 없습니다.'); return; }
      await navigator.clipboard.writeText(rawText);
      alert('결과(원본 텍스트)가 클립보드에 복사되었습니다!');
    } catch (err) {
      console.error('복사 실패:', err); alert('복사에 실패했습니다.');
    }
  });

  // 다운로드
  downloadBtns.forEach(button => {
    button.addEventListener('click', async () => {
      const format = button.getAttribute('data-format'); // pdf | docx
      if (!rawText.trim()) { alert('다운로드할 내용이 없습니다.'); return; }
      button.disabled = true;
      button.textContent = format === 'pdf' ? 'PDF 생성 중...' : 'DOCX 생성 중...';
      try {
        await downloadFile(rawText, format);
      } catch (e) {
        console.error('[DOWNLOAD fetch error]', e);
        alert(`다운로드 중 오류: ${e.message}`);
      } finally {
        button.textContent = format === 'pdf' ? '.PDF로 저장' : '.DOCX로 저장';
        button.disabled = false;
      }
    });
  });

  async function downloadFile(content, format) {
    const form = new FormData();
    form.append('content', content);
    form.append('format', format);

    const res = await fetch(`${DOWNLOAD_URL}?format=${encodeURIComponent(format)}`, {
      method: 'POST', body: form
    });
    if (!res.ok) {
      let msg = `서버 오류: ${res.status} ${res.statusText}`;
      try { const data = await res.json(); if (data && data.error) msg = data.error; } catch (_) {}
      throw new Error(msg);
    }
    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `면접_질문+답.${format}`;
    document.body.appendChild(a); a.click(); a.remove();
    window.URL.revokeObjectURL(url);
  }
});
