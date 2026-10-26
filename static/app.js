const $ = id => document.getElementById(id);
let selectedFile = null;
const fileInput = $('file');
const dropzone = $('dropzone');
const button = $('format');

function showMessage(message) {
  $('message').textContent = message;
  $('message').classList.toggle('hidden', !message);
}
function setFile(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.docx')) return showMessage('目前仅支持 .docx Word 文档。');
  if (file.size > 15 * 1024 * 1024) return showMessage('文件不能超过 15 MB。');
  selectedFile = file;
  $('filename').textContent = file.name;
  $('filesize').textContent = `${(file.size / 1024).toFixed(1)} KB`;
  $('file-card').classList.remove('hidden');
  $('report').classList.add('hidden');
  button.disabled = false;
  showMessage('');
}
dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('keydown', event => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); fileInput.click(); } });
$('choose').addEventListener('click', event => { event.stopPropagation(); fileInput.click(); });
fileInput.addEventListener('change', () => setFile(fileInput.files[0]));
['dragenter','dragover'].forEach(name => dropzone.addEventListener(name, event => { event.preventDefault(); dropzone.classList.add('drag'); }));
['dragleave','drop'].forEach(name => dropzone.addEventListener(name, event => { event.preventDefault(); dropzone.classList.remove('drag'); }));
dropzone.addEventListener('drop', event => setFile(event.dataTransfer.files[0]));
$('remove').addEventListener('click', () => { selectedFile = null; fileInput.value = ''; $('file-card').classList.add('hidden'); $('report').classList.add('hidden'); button.disabled = true; showMessage(''); });

async function post(path, body) {
  const response = await fetch(`${path}?template=${encodeURIComponent($('template').value)}`, {method:'POST', headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}, body});
  if (!response.ok) { const result = await response.json().catch(() => ({})); throw new Error(result.error || '文档处理失败'); }
  return response;
}
function renderReport(report) {
  $('report-title').textContent = report.title || '未检测到';
  $('report-count').textContent = String(report.paragraphs);
  $('report-basis').textContent = report.basis;
  $('report-status').textContent = '已完成';
  const box = $('warning-box');
  box.replaceChildren();
  if (report.warnings.length) {
    const label = document.createElement('strong'); label.textContent = '请复核'; box.append(label);
    const list = document.createElement('ul');
    report.warnings.forEach(w => { const item = document.createElement('li'); item.textContent = w; list.append(item); });
    box.append(list);
  }
  $('report').classList.remove('hidden');
}
button.addEventListener('click', async () => {
  if (!selectedFile) return;
  button.disabled = true; button.firstChild.textContent = '正在处理文档';
  showMessage('');
  try {
    const bytes = await selectedFile.arrayBuffer();
    const report = await (await post('/api/analyze', bytes)).json();
    renderReport(report);
    const response = await post('/api/format', bytes);
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url;
    a.download = `${selectedFile.name.replace(/\.docx$/i,'')}_标准格式.docx`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  } catch (error) { showMessage(error.message); }
  finally { button.disabled = false; button.firstChild.textContent = '生成标准格式文档 '; }
});
