const CLOUD_API_URL = String((window.EKARANGAN_CLOUD && window.EKARANGAN_CLOUD.apiUrl) || '').trim();
const cloudEnabled = /^https:\/\/script\.google\.com\/macros\/s\/.+\/exec(?:\?|$)/i.test(CLOUD_API_URL);
let cloudOriginalImage = '';
const localSaveEditor = saveEditor;
const localDeleteById = deleteById;
const localNewEssay = newEssay;
const localLoadEditor = loadEditor;

function cloudBadge(message, ok) {
  let el = document.getElementById('cloudStatusBadge');
  if (!el) {
    el = document.createElement('div');
    el.id = 'cloudStatusBadge';
    el.style.cssText = 'margin:0 0 10px;padding:8px 10px;border-radius:10px;font-size:12px;font-weight:800;border:1px solid #d8dee5;background:#f8fafb;color:#66717f';
    const side = document.querySelector('.teacher-side');
    if (side) side.insertBefore(el, side.querySelector('.teacher-list'));
  }
  el.textContent = message;
  if (ok === true) { el.style.background='#eef9f3';el.style.borderColor='#bfdecf';el.style.color='#176b45'; }
  else if (ok === false) { el.style.background='#fff3f3';el.style.borderColor='#efc6c6';el.style.color='#922'; }
  else { el.style.background='#f8fafb';el.style.borderColor='#d8dee5';el.style.color='#66717f'; }
}

function cloudToEssay(x) {
  return {
    id: String(x.id || ''),
    title: String(x.title || ''),
    theme: String(x.theme || ''),
    standard: String(x.standard || ''),
    easy: String(x.easy || ''),
    image: String(x.imageUrl || ''),
    active: x.active !== false
  };
}

function mergeCloudEssays(base, cloudRows) {
  const map = new Map((base || []).filter(x=>x && x.id).map(x => [x.id, JSON.parse(JSON.stringify(x))]));
  (cloudRows || []).forEach(raw => {
    const x = cloudToEssay(raw);
    if (!x.id) return;
    if (raw.active === false) { map.delete(x.id); return; }
    const old = map.get(x.id) || {};
    map.set(x.id, {...old, ...x, image: x.image || old.image || ''});
  });
  return [...map.values()];
}

async function cloudGetList() {
  const r = await fetch(CLOUD_API_URL + '?action=list&_=' + Date.now(), {cache:'no-store', redirect:'follow'});
  const text = await r.text();
  let data;
  try { data = JSON.parse(text); } catch (e) { throw new Error('Respons Apps Script bukan JSON. Semak deployment Web App.'); }
  if (!data.ok) throw new Error(data.error || 'Gagal membaca cloud.');
  return Array.isArray(data.essays) ? data.essays : [];
}

async function cloudPost(payload) {
  const r = await fetch(CLOUD_API_URL, {
    method:'POST',
    headers:{'Content-Type':'text/plain;charset=utf-8'},
    body:JSON.stringify(payload),
    redirect:'follow',
    cache:'no-store'
  });
  const text = await r.text();
  let data;
  try { data = JSON.parse(text); } catch (e) { throw new Error('Respons Apps Script bukan JSON. Semak akses Web App: Anyone.'); }
  if (!data.ok) throw new Error(data.error || 'Operasi cloud gagal.');
  return data;
}

function teacherPin() {
  return sessionStorage.getItem('ek_teacher_pin') || document.getElementById('pinInput').value || '';
}

async function cloudInit() {
  if (!cloudEnabled) {
    cloudBadge('Cloud belum aktif · data masih disimpan pada pelayar', null);
    return;
  }
  cloudBadge('Menyambung ke Google Sheets…', null);
  try {
    const rows = await cloudGetList();
    essays = mergeCloudEssays(essays, rows);
    saveData();
    if (!essays.some(x=>x.id===currentId)) currentId = essays[0]?.id || '';
    renderStudent();
    if (!document.getElementById('teacherView').classList.contains('hidden')) renderTeacherList();
    cloudBadge('Cloud aktif · Google Sheets disambungkan', true);
  } catch (err) {
    console.error(err);
    cloudBadge('Cloud tidak dapat dicapai · menggunakan data pelayar', false);
  }
}

newEssay = function() {
  cloudOriginalImage = '';
  localNewEssay();
};
loadEditor = function(id) {
  localLoadEditor(id);
  cloudOriginalImage = currentImage || '';
};

deleteById = async function(id) {
  if (!cloudEnabled) return localDeleteById(id);
  const e = essays.find(x=>x.id===id);
  if (!e) return;
  if (!confirm(`Padam “${e.title}”?`)) return;
  try {
    setStatus('Memadam dari cloud…', true);
    await cloudPost({action:'delete', pin:teacherPin(), id});
    essays = essays.filter(x=>x.id!==id);
    saveData();
    if (currentId===id) currentId=essays[0]?.id||'';
    if (currentEditId===id) newEssay();
    renderTeacherList();
    renderStudent();
    setStatus('Karangan dipadam dari cloud.', true);
    toast('Karangan dipadam untuk semua peranti.');
  } catch (err) {
    setStatus(err.message, false);
  }
};

async function cloudSaveEditor() {
  if (!cloudEnabled) return localSaveEditor();
  let title=document.getElementById('fTitle').value.trim();
  let theme=document.getElementById('fTheme').value.trim();
  let easy=document.getElementById('fEasy').value.trim();
  let standard=document.getElementById('fStandard').value.trim();
  if (!title) { setStatus('Sila masukkan tajuk.', false); return; }
  if (!easy && !standard) { setStatus('Masukkan sekurang-kurangnya satu versi karangan.', false); return; }
  if (!easy) { easy=makeEasy(standard); document.getElementById('fEasy').value=easy; document.getElementById('easyAuto').classList.remove('hidden'); }
  if (!standard) { standard=makeStandard(easy); document.getElementById('fStandard').value=standard; document.getElementById('standardAuto').classList.remove('hidden'); }

  const id=currentEditId||uid();
  const isDataImage=/^data:image\//i.test(currentImage||'');
  const imageRemoved=!!cloudOriginalImage && !currentImage;
  const essay={
    id,title,theme,easy,standard,
    imageUrl:(!isDataImage && currentImage) ? currentImage : '',
    imageRemoved
  };
  try {
    setStatus('Menyimpan ke Google Sheets…', true);
    const data=await cloudPost({action:'save', pin:teacherPin(), essay, imageData:isDataImage?currentImage:''});
    const saved=cloudToEssay(data.essay||{});
    const obj={id:saved.id||id,title:saved.title||title,theme:saved.theme||theme,easy:saved.easy||easy,standard:saved.standard||standard,image:saved.image||''};
    const i=essays.findIndex(x=>x.id===obj.id);
    if(i>=0) essays[i]=obj; else essays.push(obj);
    currentEditId=obj.id;
    currentId=obj.id;
    currentImage=obj.image;
    cloudOriginalImage=obj.image;
    saveData();
    renderTeacherList();
    renderStudent();
    document.getElementById('imagePreview').innerHTML=currentImage?`<img src="${currentImage}" alt="Pratonton">`:'<span style="color:#8b96a3">Tiada imej dipilih</span>';
    setStatus(`Disimpan ke cloud. Mudah: ${wc(easy)} perkataan · Standard: ${wc(standard)} perkataan`, true);
  } catch(err) {
    setStatus(err.message, false);
  }
}

async function authenticateTeacher() {
  const pin=document.getElementById('pinInput').value;
  const status=document.getElementById('pinStatus');
  status.className='status';status.textContent='';
  if (!cloudEnabled) {
    if(pin===DEFAULT_PIN){sessionStorage.setItem('ek_teacher_pin',pin);showTeacher();}
    else{status.textContent='PIN salah.';status.className='status show bad';}
    return;
  }
  try {
    status.textContent='Menyemak PIN…';status.className='status show';
    await cloudPost({action:'auth', pin});
    sessionStorage.setItem('ek_teacher_pin',pin);
    showTeacher();
    cloudBadge('Cloud aktif · Google Sheets disambungkan', true);
  } catch(err) {
    status.textContent=err.message;status.className='status show bad';
  }
}

const originalHandleImage = handleImage;
handleImage = function(file) {
  if (!file) return;
  if (file.size <= 2.5*1024*1024) return originalHandleImage(file);
  if (file.size > 10*1024*1024) { setStatus('Imej terlalu besar. Maksimum 10 MB sebelum pemampatan.',false); return; }
  const reader=new FileReader();
  reader.onload=()=>{
    const im=new Image();
    im.onload=()=>{
      const max=2000,scale=Math.min(1,max/Math.max(im.width,im.height));
      const c=document.createElement('canvas');c.width=Math.round(im.width*scale);c.height=Math.round(im.height*scale);
      c.getContext('2d').drawImage(im,0,0,c.width,c.height);
      currentImage=c.toDataURL('image/jpeg',0.88);
      document.getElementById('imagePreview').innerHTML=`<img src="${currentImage}" alt="Pratonton">`;
      setStatus('Imej besar telah dikecilkan secara automatik untuk cloud.',true);
    };
    im.src=reader.result;
  };
  reader.readAsDataURL(file);
};

document.getElementById('pinOk').onclick=authenticateTeacher;
document.getElementById('saveEssayBtn').onclick=cloudSaveEditor;
document.getElementById('newEssayBtn').onclick=()=>newEssay();
document.getElementById('deleteEssayBtn').onclick=()=>currentEditId?deleteById(currentEditId):newEssay();

cloudInit();
