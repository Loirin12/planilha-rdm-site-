const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.static('public'));
app.use(express.json({ limit: '10mb' }));

// 🔥 CACHE PERSISTENTE (NUNCA MAIS PERDE DADOS)
const CACHE_DIR = path.join(__dirname, 'data');
const CACHE_FILE = path.join(CACHE_DIR, 'cache.json');
let cacheGlobal = {};
let cacheTimestamp = 0;
const CACHE_TTL = 5 * 60 * 1000; // 5min

if (!fs.existsSync(CACHE_DIR)) fs.mkdirSync(CACHE_DIR);

function loadCache() {
  try {
    if (fs.existsSync(CACHE_FILE)) {
      cacheGlobal = JSON.parse(fs.readFileSync(CACHE_FILE, 'utf8'));
      cacheTimestamp = Date.now();
      console.log(`✅ Cache: ${Object.keys(cacheGlobal).length} meses`);
    }
  } catch (e) {
    console.log('Cache vazio');
  }
}

function saveCache() {
  try {
    fs.writeFileSync(CACHE_FILE, JSON.stringify(cacheGlobal, null, 2));
  } catch (e) {
    console.error('Erro cache:', e);
  }
}

loadCache(); // Carrega no startup

// ================= SUAS APIs =================
app.get('/api/meses', (req, res) => {
  res.json(Object.keys(cacheGlobal).sort());
});

app.get('/api/dias', (req, res) => {
  const { mes } = req.query;
  if (!mes || !cacheGlobal[mes]) return res.json([]);
  res.json(Object.keys(cacheGlobal[mes]).sort());
});

app.get('/api/tabela', async (req, res) => {
  const { mes, tipo } = req.query;
  
  // Cache rápido
  const agora = Date.now();
  if (cacheGlobal[mes] && agora - cacheTimestamp < CACHE_TTL) {
    return res.json(cacheGlobal[mes]);
  }
  
  try {
    const filePath = path.join(__dirname, 'data', 'planilha.xlsx');
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: 'planilha.xlsx não encontrada' });
    }
    
    const workbook = XLSX.readFile(filePath);
    const dados = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    
    // Processa por mês
    const porMes = {};
    dados.forEach(row => {
      const mesRow = row.mes || row.Mes || row['Mes'] || '';
      if (!porMes[mesRow]) porMes[mesRow] = {};
      
      const id = row.id || row.ID || row.Dia || '';
      porMes[mesRow][id] = {
        id,
        data: row.data || row.Data || '',
        pr: row.pr || row.PR || row['PR'] || '',
        emb: row.emb || row.EMB || row['EMB'] || '',
        css: row.css || row.CSS || row['CSS'] || ''
      };
    });
    
    cacheGlobal = porMes;
    cacheTimestamp = agora;
    saveCache();
    
    res.json(porMes[mes] || {});
    
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get('/api/mes-total-geral', (req, res) => {
  res.json([]); // Se precisar implementa aqui
});

// ================= SALVAR (ATUALIZA TUDO!) =================
app.post('/api/salvar', async (req, res) => {
  const { mes, dia, pr, emb, css, tipo } = req.body;
  
  console.log(`💾 ${dia} → ${pr || ''}`);
  
  try {
    // 1. Cache imediato
    if (!cacheGlobal[mes]) cacheGlobal[mes] = {};
    cacheGlobal[mes][dia] = {
      id: dia,
      pr: pr || '',
      emb: emb || '',
      css: css || ''
    };
    saveCache();
    
    // 2. Planilha backup (opcional)
    const filePath = path.join(__dirname, 'data', 'planilha.xlsx');
    if (fs.existsSync(filePath)) {
      const workbook = XLSX.readFile(filePath);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet);
      
      const linhaIndex = json.findIndex(row => 
        String(row.id || row.ID || row.Dia) === dia
      );
      
      if (linhaIndex !== -1) {
        json[linhaIndex].pr = pr;
        if (tipo === 'sig') {
          json[linhaIndex].emb = emb;
          json[linhaIndex].css = css;
        }
        
        const newSheet = XLSX.utils.json_to_sheet(json);
        workbook.Sheets[workbook.SheetNames[0]] = newSheet;
        XLSX.writeFile(workbook, filePath);
      }
    }
    
    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 app.js rodando na porta ${PORT}`);
});
