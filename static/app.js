async function q(url, opts){
  const r = await fetch(url, opts);
  if(!r.ok) throw r;
  return r.json();
}

let cacheTabela = []; // 🔥 cache para evitar múltiplos fetch pesados

/* ================= CONTROLE DE EDIÇÃO ================= */
function controlarEdicao(){
  const mes = document.getElementById('mes').value;
  const dia = document.getElementById('dia').value;

  const campos = [
    document.getElementById('pr'),
    document.getElementById('emb'),
    document.getElementById('css')
  ].filter(Boolean);

  const btnSalvar = document.getElementById('salvar');

  if(mes === 'TOTAL GERAL' || !dia){
    campos.forEach(c => {
      c.value = '';
      c.disabled = true;
    });
    btnSalvar.disabled = true;
    return;
  }

  campos.forEach(c => c.disabled = false);
  btnSalvar.disabled = false;
}

/* ================= MESES (SEM LOOP) ================= */
async function carregarMeses(){
  const meses = await q('/api/meses');
  const sel = document.getElementById('mes');

  sel.innerHTML = '';

  const usados = new Set();

  meses.forEach(m => {
    const nome = m.trim();
    if (usados.has(nome)) return;

    usados.add(nome);

    const opt = document.createElement('option');
    opt.value = nome;
    opt.textContent = nome;
    sel.appendChild(opt);
  });

  // 🔥 NÃO usar dispatchEvent (causa loop e lentidão)
  if (sel.options.length > 0) {
    sel.selectedIndex = 0;
    await atualizarDias();
    await carregarTabela();
  }
}

/* ================= DIAS ================= */
async function atualizarDias(){
  const mes = document.getElementById('mes').value;
  const sel = document.getElementById('dia');

  sel.innerHTML = '<option value="">Selecione</option>';

  if(!mes || mes === 'TOTAL GERAL'){
    controlarEdicao();
    return;
  }

  const dias = await q(`/api/dias?mes=${encodeURIComponent(mes)}`);

  dias.forEach(d=>{
    const o = document.createElement('option');
    o.value = d;
    o.textContent = d;
    sel.appendChild(o);
  });

  controlarEdicao();
}

/* ================= TABELA (COM CACHE - MUITO MAIS RÁPIDO) ================= */
async function carregarTabela(){
  const mes = document.getElementById('mes').value;
  const tbody = document.getElementById('tbody');
  tbody.innerHTML = '';

  if(!mes) return;

  let url;
  if(mes === 'TOTAL GERAL'){
    url = `/api/mes-total-geral?tipo=${TIPO}`;
  } else {
    url = `/api/tabela?mes=${encodeURIComponent(mes)}&tipo=${TIPO}`;
  }

  // 🔥 Apenas 1 fetch pesado (Excel)
  cacheTabela = await q(url);

  renderTabela();
}

function renderTabela(){
  const tbody = document.getElementById('tbody');
  tbody.innerHTML = '';

  cacheTabela.forEach(r=>{
    const tr = document.createElement('tr');

    tr.innerHTML = `
      <td>${r.id || ''}</td>
      <td>${r.data || ''}</td>
      <td>${r.pr || ''}</td>
      ${TIPO === 'sig' ? `<td>${r.emb || ''}</td><td>${r.css || ''}</td>` : ''}
    `;

    tr.style.pointerEvents = 'none';
    tbody.appendChild(tr);
  });
}

/* ================= CARREGAR DIA (SEM NOVO FETCH!) ================= */
function carregarDia(){
  const mes = document.getElementById('mes').value;
  const dia = document.getElementById('dia').value;

  if(!mes || !dia || mes === 'TOTAL GERAL'){
    controlarEdicao();
    return;
  }

  // 🔥 usa cache em vez de chamar API de novo (grande ganho de performance)
  const linha = cacheTabela.find(l => String(l.id) === String(dia));

  document.getElementById('pr').value = linha?.pr || '';

  if(TIPO === 'sig'){
    document.getElementById('emb').value = linha?.emb || '';
    document.getElementById('css').value = linha?.css || '';
  }

  controlarEdicao();
}

/* ================= LOAD (MAIS RÁPIDO QUE window.load) ================= */
document.addEventListener('DOMContentLoaded', async ()=>{
  await carregarMeses();
  controlarEdicao();

  document.getElementById('mes').addEventListener('change', async ()=>{
    await atualizarDias();
    await carregarTabela(); // 1 único fetch pesado
  });

  document.getElementById('dia').addEventListener('change', carregarDia);

  document.getElementById('salvar').addEventListener('click', async ()=>{
    const mes = document.getElementById('mes').value;
    const dia = document.getElementById('dia').value;

    if(mes === 'TOTAL GERAL' || !dia) return;

    const payload = {
      mes,
      dia,
      pr: document.getElementById('pr').value,
      tipo: TIPO
    };

    if(TIPO === 'sig'){
      payload.emb = document.getElementById('emb').value;
      payload.css = document.getElementById('css').value;
    }

    const res = await fetch('/api/salvar',{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify(payload)
    });

    if(res.ok){
      alert('Salvo com sucesso!');
      await carregarTabela(); // atualiza cache
      carregarDia(); // sem novo fetch
    } else {
      alert('Erro ao salvar');
    }
  });
});
