async function q(url, opts){
  try{
    const r = await fetch(url, opts);
    if(!r.ok) throw new Error(`Erro HTTP: ${r.status}`);
    return await r.json();
  }catch(e){
    console.error('Erro na requisição:', url, e);
    return [];
  }
}

/* ================= CONTROLE DE EDIÇÃO ================= */
function controlarEdicao(){
  const mes = document.getElementById('mes')?.value;
  const dia = document.getElementById('dia')?.value;

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
    if(btnSalvar) btnSalvar.disabled = true;
    return;
  }

  campos.forEach(c => c.disabled = false);
  if(btnSalvar) btnSalvar.disabled = false;
}

/* ================= MESES (ANTI-LOOP) ================= */
let mesesCarregados = false;

async function carregarMeses(){
  if(mesesCarregados) return; // 🔥 evita loop de carregamento
  mesesCarregados = true;

  const meses = await q('/api/meses');
  const sel = document.getElementById('mes');

  if(!sel) return;

  sel.innerHTML = '';

  const usados = new Set();

  meses.forEach(m => {
    const nome = String(m).trim();
    if(!nome || usados.has(nome)) return;

    usados.add(nome);

    const opt = document.createElement('option');
    opt.value = nome;
    opt.textContent = nome;
    sel.appendChild(opt);
  });

  // seleciona primeiro mês SEM disparar loop infinito
  if(sel.options.length > 0){
    sel.selectedIndex = 0;
    await atualizarDias();
    await carregarTabela();
  }
}

/* ================= DIAS ================= */
async function atualizarDias(){
  const mes = document.getElementById('mes')?.value;
  const sel = document.getElementById('dia');

  if(!sel) return;

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

/* ================= TABELA ================= */
let carregandoTabela = false;

async function carregarTabela(){
  if(carregandoTabela) return; // 🔥 trava anti-loop
  carregandoTabela = true;

  try{
    const mes = document.getElementById('mes')?.value;
    const tbody = document.getElementById('tbody');

    if(!tbody || !mes){
      carregandoTabela = false;
      return;
    }

    tbody.innerHTML = '';

    let url;
    if(mes === 'TOTAL GERAL'){
      url = `/api/mes-total-geral?tipo=${TIPO}`;
    } else {
      url = `/api/tabela?mes=${encodeURIComponent(mes)}&tipo=${TIPO}`;
    }

    const rows = await q(url);

    rows.forEach(r=>{
      const tr = document.createElement('tr');

      tr.innerHTML = `
        <td>${r.id ?? ''}</td>
        <td>${r.data ?? ''}</td>
        <td>${r.pr ?? ''}</td>
        ${TIPO === 'sig' 
          ? `<td>${r.emb ?? ''}</td><td>${r.css ?? ''}</td>` 
          : ''}
      `;

      tr.style.pointerEvents = 'none';
      tbody.appendChild(tr);
    });

  }finally{
    carregandoTabela = false;
  }
}

/* ================= CARREGAR DIA ================= */
let carregandoDia = false;

async function carregarDia(){
  if(carregandoDia) return; // 🔥 evita loop
  carregandoDia = true;

  try{
    const mes = document.getElementById('mes')?.value;
    const dia = document.getElementById('dia')?.value;

    if(!mes || !dia || mes === 'TOTAL GERAL'){
      controlarEdicao();
      return;
    }

    const dados = await q(`/api/tabela?mes=${encodeURIComponent(mes)}&tipo=${TIPO}`);
    const linha = dados.find(l => String(l.id) === String(dia));

    const pr = document.getElementById('pr');
    if(pr) pr.value = linha?.pr ?? '';

    if(TIPO === 'sig'){
      const emb = document.getElementById('emb');
      const css = document.getElementById('css');
      if(emb) emb.value = linha?.emb ?? '';
      if(css) css.value = linha?.css ?? '';
    }

    controlarEdicao();
  }finally{
    carregandoDia = false;
  }
}

/* ================= INIT ================= */
window.addEventListener('DOMContentLoaded', async ()=>{
  const mesEl = document.getElementById('mes');
  const diaEl = document.getElementById('dia');
  const salvarEl = document.getElementById('salvar');

  await carregarMeses();
  controlarEdicao();

  mesEl?.addEventListener('change', async ()=>{
    await atualizarDias();
    await carregarTabela();
  });

  diaEl?.addEventListener('change', carregarDia);

  salvarEl?.addEventListener('click', async ()=>{
    const mes = mesEl?.value;
    const dia = diaEl?.value;

    if(!mes || mes === 'TOTAL GERAL' || !dia) return;

    const payload = {
      mes,
      dia,
      pr: document.getElementById('pr')?.value || '',
      tipo: TIPO
    };

    if(TIPO === 'sig'){
      payload.emb = document.getElementById('emb')?.value || '';
      payload.css = document.getElementById('css')?.value || '';
    }

    const res = await fetch('/api/salvar',{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify(payload)
    });

    if(res.ok){
      alert('Salvo com sucesso!');
      await carregarTabela();
      await carregarDia();
    } else {
      alert('Erro ao salvar');
    }
  });
});
