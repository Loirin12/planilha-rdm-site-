const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.static('public'));
app.use(express.json({ limit: '10mb' }));

// 🔥 CACHE PERSISTENTE (sobrevive restarts)
const CACHE_DIR = path.join(__dirname, 'data');
const CACHE_FILE = path.join(CACHE_DIR, 'cache.json');
let cacheGlobal = {};
let cacheTimestamp = 0;
const CACHE_TTL = 5 * 60 * 1000; // 5min

// 🛠️ Garante pasta data existe
if (!fs.existsSync(CACHE_DIR)) fs.mkdirSync(CACHE_DIR);

// 💾 Carrega cache do disco
function loadCache() {
  try {
    if (fs.existsSync(CACHE_FILE)) {
      cacheGlobal = JSON.parse(fs.readFileSync(CACHE_FILE, 'utf8'));
      cacheTimestamp = Date.now();
      console.log(`✅ Cache carregado: ${Object.keys(cacheGlobal).length} meses`);
    }
  } catch (e) {
    console.log('Cache corrompido, iniciando vazio');
  }
}

// 💾 Salva cache no disco
function saveCache() {
  try {
    fs.writeFileSync(CACHE_FILE, JSON.stringify(cacheGlobal, null, 2));
    console.log('💾 Cache salvo no disco');
  } catch (e) {
    console.error('Erro salvando cache:', e);
  }
}

// Inicializa cache
loadCache();

// ================= API MESES =================
app.get('/api/meses', (req, res) => {
  const meses = Object.keys(cacheGlobal).sort();
  res.json(meses);
});

// ================= API DIAS =================
app.get('/api/dias', (req, res) => {
  const { mes } = req.query;
  if (!mes || !cacheGlobal[mes]) {
    return res.json([]);
  }
  const dias = Object.keys(cacheGlobal[mes]).sort();
  res.json(dias);
});

// ================= API TABELA (SEU CÓDIGO!) =================
app.get('/api/tabela', async (req, res) => {
  const { mes, tipo } = req.query;
  
  // Cache válido?
  const agora = Date.now();
  if (cacheGlobal[mes] && agora - cacheTimestamp < CACHE_TTL) {
    console.log(`📦 Cache hit: ${mes}`);
    return res.json(cacheGlobal[mes]);
  }
  
  console.log(`🔄 Recarregando ${mes}...`);
  
  try {
    // Lê planilha master
    const filePath = path.join(__dirname, 'data', 'planilha.xlsx');
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: 'planilha.xlsx não encontrada' });
    }
    
    const workbook = XLSX.readFile(filePath);
    const dados = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    
    // 🔥 Processa por mês (igual seu frontend espera)
    const porMes = {};
    dados.forEach(row => {
      const mesRow = row.mes || row.Mes || '';
      if (!porMes[mesRow]) porMes[mesRow] = {};
      
      const id = row.id || row.ID || row.Dia || '';
      porMes[mesRow][id] = {
        id,
        data: row.data || row.Data || '',
        pr: row.pr || row.PR || row['PR'] || '',
        emb: row.emb || row.EMB || '',
        css: row.css || row.CSS || ''
      };
    });
    
    // Salva cache
    cacheGlobal = porMes;
    cacheTimestamp = agora;
    saveCache();
    
    console.log(`✅ ${mes} processado: ${Object.keys(porMes[mes] || {}).length} dias`);
    res.json(porMes[mes] || {});
    
  } catch (error) {
    console.error('Erro planilha:', error);
    res.status(500).json({ error: 'Erro processando planilha' });
  }
});

// ================= API TOTAL GERAL =================
app.get('/api/mes-total-geral', (req, res) => {
  const { tipo } = req.query;
  // Lógica do total geral aqui (se precisar)
  res.json([]);
});

// ================= SALVAR (ATUALIZA PLANILHA!) =================
app.post('/api/salvar', async (req, res) => {
  const { mes, dia, pr, emb, css, tipo } = req.body;
  
  console.log(`💾 Salvando ${dia} em ${mes}`);
  
  try {
    // 1️⃣ Atualiza cache imediatamente
    if (!cacheGlobal[mes]) cacheGlobal[mes] = {};
    cacheGlobal[mes][dia] = {
      id: dia,
      pr: pr || '',
      emb: emb || '',
      css: css || ''
    };
    saveCache();
    
    // 2️⃣ Atualiza planilha.xlsx (OPCIONAL - para backup)
    const filePath = path.join(__dirname, 'data', 'planilha.xlsx');
    if (fs.existsSync(filePath)) {
      const workbook = XLSX.readFile(filePath);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet);
      
      // Encontra e atualiza linha
      const linhaIndex = json.findIndex(row => String(row.id || row.ID || row.Dia) === dia);
      if (linhaIndex !== -1) {
        json[linhaIndex].pr = pr;
        if (tipo === 'sig') {
          json[linhaIndex].emb = emb;
          json[linhaIndex].css = css;
        }
        
        // Reescreve planilha
        const newSheet = XLSX.utils.json_to_sheet(json);
        workbook.Sheets[workbook.SheetNames[0]] = newSheet;
        XLSX.writeFile(workbook, filePath);
        console.log('📄 Planilha.xlsx atualizada');
      }
    }
    
    res.json({ success: true });
  } catch (error) {
    console.error('Erro salvando:', error);
    res.status(500).json({ error: 'Erro ao salvar' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 RDM API em ${PORT}`);
  console.log(`📊 Cache: ${Object.keys(cacheGlobal).length} meses carregados`);
});
