let sciPath = null;
let zipPath = null;
let currentTaskId = null;

// --- MAIN FILE UPLOADS ---
function setupDrop(dropId, fileId, callback) {
    const drop = document.getElementById(dropId);
    const input = document.getElementById(fileId);

    drop.addEventListener('click', () => input.click());
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('dragover'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
    
    drop.addEventListener('drop', e => {
        e.preventDefault();
        drop.classList.remove('dragover');
        if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0], drop, callback);
    });
    
    input.addEventListener('change', e => {
        if (e.target.files.length) handleFile(e.target.files[0], drop, callback);
    });
}

async function handleFile(file, dropEl, callback) {
    const formData = new FormData();
    formData.append('file', file);
    
    dropEl.classList.add('loading');
    const res = await fetch(dropEl.id.includes('sci') ? '/upload/sci' : '/upload/zip', { method: 'POST', body: formData });
    const data = await res.json();
    
    dropEl.classList.remove('loading');
    if (data.path) {
        dropEl.classList.add('success');
        dropEl.querySelector('p').innerText = file.name;
        callback(data.path);
    }
}

setupDrop('drop-sci', 'file-sci', path => { sciPath = path; checkReady(); });
setupDrop('drop-zip', 'file-zip', path => { zipPath = path; checkReady(); });

function checkReady() {
    document.getElementById('btn-process').disabled = !(sciPath && zipPath);
}

// --- PROCESSING ---
document.getElementById('btn-process').addEventListener('click', async () => {
    const logBox = document.getElementById('log-box');
    const btnDown = document.getElementById('btn-download');
    logBox.innerHTML = '';
    btnDown.classList.add('hidden');
    document.getElementById('btn-process').disabled = true;

    const res = await fetch('/processar', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sci_path: sciPath, zip_path: zipPath })
    });
    const task = await res.json();
    currentTaskId = task.task_id;

    const poll = setInterval(async () => {
        const statusRes = await fetch(`/status/${currentTaskId}`);
        const statusData = await statusRes.json();

        logBox.innerHTML = statusData.log.map(l => `<div>${l}</div>`).join('');
        logBox.scrollTop = logBox.scrollHeight;

        if (statusData.status === 'DONE' || statusData.status === 'ERROR') {
            clearInterval(poll);
            document.getElementById('btn-process').disabled = false;
            if (statusData.status === 'DONE') btnDown.classList.remove('hidden');
        }
    }, 1000);
});

function downloadFile() {
    window.location.href = `/download/${currentTaskId}`;
}

// --- REFERENCE DATA MANAGEMENT ---
['funcionarios', 'fornecedores', 'cheques'].forEach(type => {
    const drop = document.getElementById(`drop-${type}`);
    const input = document.getElementById(`file-${type}`);

    drop.addEventListener('click', () => input.click());
    drop.addEventListener('dragover', e => e.preventDefault());
    drop.addEventListener('drop', async e => {
        e.preventDefault();
        if(!e.dataTransfer.files.length) return;
        const formData = new FormData();
        formData.append('file', e.dataTransfer.files[0]);
        await fetch(`/data/${type}`, { method: 'POST', body: formData });
        loadRefData();
    });
    input.addEventListener('change', async e => {
        if(!e.target.files.length) return;
        const formData = new FormData();
        formData.append('file', e.target.files[0]);
        await fetch(`/data/${type}`, { method: 'POST', body: formData });
        loadRefData();
    });
});

async function loadRefData() {
    ['funcionarios', 'fornecedores', 'cheques'].forEach(async type => {
        const res = await fetch(`/data/${type}`);
        const data = await res.json();
        document.getElementById(`count-${type}`).innerText = `${data.length} registros carregados`;
    });
}

loadRefData();