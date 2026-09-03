const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const XLSX = require('xlsx');
const fs = require('fs');

const app = express();
const db = new sqlite3.Database('./voting.db');

app.use(express.json());
app.use(express.static(path.join(__dirname)));

// ═══ ПУТЬ К РЕЕСТРУ ═══
// На вашем ПК оставьте старый путь, на сервере положите файл рядом с server.js
const REGISTRY_PATH = process.env.REGISTRY_PATH || './Сводный_реестр_сверка.xlsx';

// ═══ КЭШ РЕЕСТРА ═══
let registryCache = null;
let registryCacheTime = 0;
const CACHE_TTL = 60000;

function loadRegistry() {
    const now = Date.now();
    if (registryCache && (now - registryCacheTime) < CACHE_TTL) {
        return registryCache;
    }
    try {
        if (!fs.existsSync(REGISTRY_PATH)) {
            console.error(`❌ Файл реестра не найден: ${REGISTRY_PATH}`);
            return [];
        }
        const workbook = XLSX.readFile(REGISTRY_PATH);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        registryCache = XLSX.utils.sheet_to_json(worksheet);
        registryCacheTime = now;
        console.log(`[${new Date().toLocaleTimeString()}] Реестр загружен: ${registryCache.length} записей`);
        return registryCache;
    } catch (error) {
        console.error("❌ Ошибка чтения реестра:", error.message);
        return [];
    }
}

// ═══ ИНИЦИАЛИЗАЦИЯ БД ═══
db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS voters (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email TEXT UNIQUE,
        plot TEXT,
        voted_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS votes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        q_smeta TEXT,
        snt_status TEXT,
        board_eval TEXT,
        priorities TEXT,
        comments TEXT
    )`);
});

// ═══ ПРОВЕРКА ПО РЕЕСТРУ ═══
function verifyAgainstRegistry(email, plot) {
    const registryData = loadRegistry();
    const cleanEmail = email.trim().toLowerCase();
    const cleanPlot  = plot.trim().toLowerCase();

    return registryData.some(row => {
        let rowEmail = '';
        let rowPlot  = '';
        for (let key in row) {
            const normalizedKey = key.trim().toLowerCase();
            if (normalizedKey === 'e-mail' || normalizedKey === 'email' || normalizedKey === 'почта' || normalizedKey === 'e-mail адрес') {
                rowEmail = String(row[key] || '').trim().toLowerCase();
            }
            if (normalizedKey.includes('участ') || normalizedKey === '№' || normalizedKey === 'номер' || normalizedKey === '№ уч.' || normalizedKey === 'уч.' || normalizedKey === 'участок') {
                rowPlot = String(row[key] || '').trim().toLowerCase();
            }
        }
        return rowEmail === cleanEmail && rowPlot === cleanPlot;
    });
}

// ═══ ПРИЁМ ГОЛОСА ═══
app.post('/api/vote', (req, res) => {
    const { email, plot, q_smeta, snt_status, board_eval, priorities, comments } = req.body;

    if (!email || !plot || !q_smeta || !snt_status || !board_eval) {
        return res.status(400).json({ error: 'Заполните все обязательные поля' });
    }

    const isAllowed = verifyAgainstRegistry(email, plot);
    if (!isAllowed) {
        return res.status(403).json({ 
            error: 'Указанный E-mail и номер участка не найдены в Сводном реестре домовладений (или не совпадают).' 
        });
    }

    db.get('SELECT id FROM voters WHERE email = ?', [email.trim().toLowerCase()], (err, row) => {
        if (err) return res.status(500).json({ error: 'Ошибка сервера' });
        if (row) return res.status(400).json({ error: 'Вы уже принимали участие в данном голосовании' });

        db.run('INSERT INTO voters (email, plot) VALUES (?, ?)', [email.trim().toLowerCase(), plot.trim()], function(err) {
            if (err) return res.status(500).json({ error: 'Ошибка записи участника' });

            const prioritiesString = Array.isArray(priorities) ? priorities.join(', ') : '';
            db.run(
                'INSERT INTO votes (q_smeta, snt_status, board_eval, priorities, comments) VALUES (?, ?, ?, ?, ?)',
                [q_smeta, snt_status, board_eval, prioritiesString, comments || ''],
                (err) => {
                    if (err) return res.status(500).json({ error: 'Ошибка сохранения голоса' });
                    res.json({ success: true, message: 'Ваш голос успешно сверен со Сводным реестром и анонимно учтен!' });
                }
            );
        });
    });
});

// ═══ ГЛАВНАЯ ═══
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// ═══ АДМИНКА ═══
app.post('/api/admin/data', (req, res) => {
    const { password } = req.body;
    if (password !== 'admin123') {
        return res.status(401).json({ error: 'Неверный пароль администратора' });
    }

    db.all('SELECT email, plot, voted_at FROM voters ORDER BY voted_at DESC', [], (err, voters) => {
        if (err) return res.status(500).json({ error: 'Ошибка базы данных' });

        db.all('SELECT q_smeta, snt_status, board_eval, priorities, comments FROM votes', [], (err, votes) => {
            if (err) return res.status(500).json({ error: 'Ошибка базы данных' });
            res.json({ voters, votes });
        });
    });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`✅ Сервер запущен на порту ${PORT}`);
    console.log(`📂 Реестр: ${REGISTRY_PATH}`);
});
