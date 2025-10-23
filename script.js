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
  const resultArea = document.getElementById('result-area');   // HTML(마크다운 렌더)
  const rawArea = document.getElementById('raw-area');         // 원본 텍스트
  const errorContainer = document.getElementById('error-container');
  const errorMessage = document.getElementById('error-message');
  const copyBtn = document.getElementById('copy-btn');
  const downloadBtns = document.querySelectorAll('.download-btn');
  const hwpInfoBtn = document.getElementById('hwp-info-btn');

  // ✅ 항상 백엔드로 보내도록 BASE 고정 (혼합콘텐츠/다른 포트 문제 방지)
  //const API_BASE = window.API_BASE || 'http://127.0.0.1:5000';
  const API_BASE = window.API_BASE || 'https://port-0-interviewpro-mh3iopw4cf627816.sel3.cloudtype.app';
  const GENERATE_URL = `${API_BASE}/generate`;
  const DOWNLOAD_URL = `${API_BASE}/download`;

  // 상태
  let selectedFile = null;    // 드래그&드롭 또는 input 선택 파일
  let rawText = '';           // 스트리밍 받은 원본 텍스트(마크다운 포함)

  // 마크다운 옵션
  marked.use({ breaks: true, gfm: true });

  // ========== 드래그&드롭 ==========
  ['dragenter','dragover'].forEach(ev => {
    dropzone.addEventListener(ev, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('dragover');
      fileLabelText.textContent = '여기에 놓으면 업로드됩니다!';
    });
  });

  ['dragleave','drop'].forEach(ev => {
    dropzone.addEventListener(ev, (e) => {
      e.preventDefault();
      e.stopPropagation();
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

  // input 파일 선택
  pdfFileInput.addEventListener('change', () => {
    const file = pdfFileInput.files[0];
    if (file) {
      if (!file.type.startsWith('application/pdf') && !file.name.toLowerCase().endsWith('.pdf')) {
        alert('PDF 파일만 선택해주세요.');
        pdfFileInput.value = '';
        selectedFile = null;
        generateBtn.disabled = true;
        return;
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

  // 결과/에러 초기화
  function hideResults() {
    resultContainer.classList.add('hidden');
    errorContainer.classList.add('hidden');
    resultArea.innerHTML = '';
    rawArea.textContent = '';
    rawText = '';
    setActionButtonsEnabled(false);
  }

  // 액션 버튼 활성화
  function setActionButtonsEnabled(enabled) {
    copyBtn.disabled = !enabled;
    downloadBtns.forEach(b => (b.disabled = !enabled));
  }

  // 로딩 상태
  function setLoadingState(isLoading) {
    if (isLoading) {
      generateBtn.disabled = true;
      generateBtn.textContent = 'AI 분석 중...';
      loadingDiv.classList.remove('hidden');
    } else {
      generateBtn.disabled = false;
      generateBtn.textContent = '✨ 질문 생성 시작';
      loadingDiv.classList.add('hidden');
    }
  }

  // 에러 표시
  function displayError(message) {
    errorMessage.textContent = message;
    errorContainer.classList.remove('hidden');
  }

  // 마크다운 렌더(안전)
  function renderMarkdownToResult(text) {
    const html = DOMPurify.sanitize(marked.parse(text || ''));
    resultArea.innerHTML = html;
  }

  // 자동 스크롤
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

      // 네트워크 레벨 실패 시 response 자체가 생성되지 않아 catch로 떨어짐
      if (!response.ok) {
        const msg = `서버 오류: ${response.status} ${response.statusText}`;
        throw new Error(msg);
      }

      resultContainer.classList.remove('hidden');
      resultArea.innerHTML = '';
      rawText = '';
      rawArea.textContent = '';

      // 스트리밍 처리
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
      // fetch가 아예 실패하면 대부분 CORS/오리진/혼합콘텐츠/주소문제
      displayError(err.message || '요청에 실패했습니다(Failed to fetch). 백엔드 주소 또는 CORS/프로토콜을 확인하세요.');
    } finally {
      setLoadingState(false);
    }
  });

  // 결과 복사(원본 마크다운 기준)
  copyBtn.addEventListener('click', async () => {
    try {
      if (!rawText.trim()) {
        alert('복사할 결과가 없습니다.');
        return;
      }
      await navigator.clipboard.writeText(rawText);
      alert('결과(원본 텍스트)가 클립보드에 복사되었습니다!');
    } catch (err) {
      console.error('복사 실패:', err);
      alert('복사에 실패했습니다.');
    }
  });

  // 다운로드 버튼
  downloadBtns.forEach(button => {
    button.addEventListener('click', async () => {
      const format = button.getAttribute('data-format'); // pdf | docx
      if (!rawText.trim()) {
        alert('다운로드할 내용이 없습니다. 먼저 결과를 생성하세요.');
        return;
      }
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

  // HWP 안내
  hwpInfoBtn.addEventListener('click', () => {
    alert(
      'HWP 파일 저장은 직접 복사-붙여넣기를 이용해주세요.\n\n' +
      '1) [결과 복사하기] 버튼 클릭\n' +
      '2) 한컴오피스(한글) 새 문서 열기\n' +
      '3) 붙여넣기(Ctrl+V)\n' +
      '4) 필요 시 서식/줄 간격 정돈 후 저장'
    );
  });

  // 파일 생성 요청 → 브라우저 다운로드
  async function downloadFile(content, format) {
    const form = new FormData();
    form.append('content', content);
    form.append('format', format);

    const res = await fetch(`${DOWNLOAD_URL}?format=${encodeURIComponent(format)}`, {
      method: 'POST',
      body: form
    });

    if (!res.ok) {
      let msg = `서버 오류: ${res.status} ${res.statusText}`;
      try {
        const data = await res.json();
        if (data && data.error) msg = data.error;
      } catch (_) {}
      throw new Error(msg);
    }

    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `AI_면접_질문.${format}`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
  }
});


