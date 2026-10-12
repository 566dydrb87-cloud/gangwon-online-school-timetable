/**
 * 강원온라인학교 시간표 시스템 백엔드 API 서버
 * - 포트 3000 (또는 환경변수 PORT)
 * - 시간표 데이터 영구 저장: data/timetable-db.json
 * - 정적 웹 페이지 서빙: index.html 및 모든 정적 자산
 */

const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const DB_FILE = path.join(__dirname, 'data', 'timetable-db.json');

// 미들웨어 설정
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// 정적 파일 서빙
app.use(express.static(__dirname));

// DB 파일 보장
function ensureDbFile() {
  const dir = path.dirname(DB_FILE);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
  if (!fs.existsSync(DB_FILE)) {
    fs.writeFileSync(DB_FILE, JSON.stringify({
      periods: [],
      students: [],
      offeredSubjects: [],
      subjectColorMap: {},
      lastUpdated: new Date().toISOString()
    }, null, 2), 'utf-8');
  }
}

// 1. 상태 확인 (Health check)
app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    serverTime: new Date().toISOString(),
    service: '강원온라인학교 시간표 백엔드 서버'
  });
});

// 2. 시간표 데이터 조회 API (GET)
app.get('/api/timetable', (req, res) => {
  try {
    ensureDbFile();
    const data = fs.readFileSync(DB_FILE, 'utf-8');
    const parsed = JSON.parse(data);
    res.json({ success: true, data: parsed });
  } catch (err) {
    console.error('시간표 데이터 조회 실패:', err);
    res.status(500).json({ success: false, message: '데이터 조회 실패', error: err.message });
  }
});

// 3. 시간표 데이터 저장 API (POST)
app.post('/api/timetable', (req, res) => {
  try {
    ensureDbFile();
    const payload = req.body;
    if (!payload || typeof payload !== 'object') {
      return res.status(400).json({ success: false, message: '유효하지 않은 데이터입니다.' });
    }

    const dataToSave = {
      ...payload,
      lastUpdated: new Date().toISOString()
    };

    fs.writeFileSync(DB_FILE, JSON.stringify(dataToSave, null, 2), 'utf-8');
    console.log(`[${new Date().toLocaleTimeString()}] 시간표 데이터가 백엔드에 영구 저장되었습니다.`);
    res.json({ success: true, message: '저장 완료', lastUpdated: dataToSave.lastUpdated });
  } catch (err) {
    console.error('시간표 데이터 저장 실패:', err);
    res.status(500).json({ success: false, message: '데이터 저장 실패', error: err.message });
  }
});

// 루트 요청 시 index.html 전송
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// 서버 시작
app.listen(PORT, () => {
  console.log('====================================================');
  console.log(`🚀 강원온라인학교 시간표 백엔드 서버가 시작되었습니다!`);
  console.log(`📡 접속 주소: http://localhost:${PORT}`);
  console.log(`💾 데이터베이스 파일: ${DB_FILE}`);
  console.log('====================================================');
});
