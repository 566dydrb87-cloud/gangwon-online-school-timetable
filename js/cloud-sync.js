/**
 * Cloud & Backend Sync Service
 * 지원 모드:
 * 1. Node.js Express 백엔드 API (REST: /api/timetable)
 * 2. Firebase Cloud Firestore 실시간 클라우드 DB
 * 3. 브라우저 LocalStorage 자동 백업 (오프라인 보호)
 */

class CloudSyncService {
  constructor() {
    this.serverUrl = window.location.origin.includes('http') && !window.location.origin.includes('github.io')
      ? window.location.origin
      : 'http://localhost:3000';
    this.activeBackend = 'none'; // 'server', 'firebase', 'none'
    this.firebaseConfigKey = 'gw_timetable_firebase_config_v1';
    this.firebaseApp = null;
    this.firestoreDb = null;
    this.isSyncing = false;

    this.init();
  }

  async init() {
    // 1. Firebase 설정 확인 및 초기화 시도
    this.checkFirebaseConfig();

    // 2. Node.js 백엔드 서버 감지 시도
    await this.detectNodeServer();
  }

  // 1. Node.js 백엔드 서버 감지
  async detectNodeServer() {
    try {
      const res = await fetch(`${this.serverUrl}/api/health`, { method: 'GET', signal: AbortSignal.timeout(1500) });
      if (res.ok) {
        const data = await res.json();
        if (data.status === 'ok') {
          this.activeBackend = 'server';
          this.updateUiBadge('server');
          console.log('✅ Node.js 백엔드 서버와 연결되었습니다:', this.serverUrl);
          return true;
        }
      }
    } catch (e) {
      // 로컬 서버가 실행 중이 아님 (정상 동작)
    }

    if (this.firestoreDb) {
      this.activeBackend = 'firebase';
      this.updateUiBadge('firebase');
    } else {
      this.activeBackend = 'none';
      this.updateUiBadge('none');
    }
    return false;
  }

  // 2. Firebase 설정 로드 및 초기화
  checkFirebaseConfig() {
    try {
      const savedConfig = localStorage.getItem(this.firebaseConfigKey);
      if (savedConfig && window.firebase) {
        const config = JSON.parse(savedConfig);
        if (config.apiKey && config.projectId) {
          if (!firebase.apps.length) {
            this.firebaseApp = firebase.initializeApp(config);
          } else {
            this.firebaseApp = firebase.app();
          }
          this.firestoreDb = firebase.firestore();
          this.activeBackend = 'firebase';
          this.updateUiBadge('firebase');
          console.log('✅ Firebase Firestore 클라우드 DB와 연결되었습니다:', config.projectId);
          this.subscribeFirebaseChanges();
        }
      }
    } catch (e) {
      console.warn('Firebase 초기화 오류:', e);
    }
  }

  // Firebase 실시간 리스너
  subscribeFirebaseChanges() {
    if (!this.firestoreDb) return;
    try {
      this.firestoreDb.collection('timetables').doc('gw_online_school')
        .onSnapshot((doc) => {
          if (doc.exists && !this.isSyncing) {
            const remoteData = doc.data();
            if (remoteData && remoteData.lastUpdated) {
              const localSaved = localStorage.getItem(window.app?.storageKey);
              const localData = localSaved ? JSON.parse(localSaved) : null;
              // 원격 데이터가 더 최신인 경우 자동 반영
              if (!localData || !localData.lastUpdated || new Date(remoteData.lastUpdated) > new Date(localData.lastUpdated)) {
                console.log('🔄 다른 기기에서 수정된 최신 시간표를 수신했습니다.');
                window.app?.applyExternalData(remoteData, '클라우드 실시간 동기화로 최신 데이터가 반영되었습니다.');
              }
            }
          }
        }, (err) => {
          console.warn('Firestore 리스너 에러:', err);
        });
    } catch (e) {
      console.error('Firebase 구독 설정 실패:', e);
    }
  }

  // 3. 데이터 영구 저장 (백엔드 및 클라우드로 전송)
  async saveToBackend(data) {
    this.isSyncing = true;
    let saved = false;

    // A. Node.js 백엔드 서버 저장
    if (this.activeBackend === 'server') {
      try {
        const res = await fetch(`${this.serverUrl}/api/timetable`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(data)
        });
        if (res.ok) {
          saved = true;
          this.updateStatusText('백엔드 서버에 저장됨');
        }
      } catch (e) {
        console.warn('Node 서버 저장 실패:', e);
      }
    }

    // B. Firebase Firestore 클라우드 저장
    if (this.firestoreDb) {
      try {
        await this.firestoreDb.collection('timetables').doc('gw_online_school').set({
          ...data,
          lastUpdated: new Date().toISOString()
        });
        saved = true;
        this.updateStatusText('클라우드 DB(Firestore)에 실시간 보존됨');
      } catch (e) {
        console.warn('Firestore 저장 실패:', e);
      }
    }

    setTimeout(() => { this.isSyncing = false; }, 500);
    return saved;
  }

  // 4. 데이터 불러오기
  async loadFromBackend() {
    if (this.activeBackend === 'server') {
      try {
        const res = await fetch(`${this.serverUrl}/api/timetable`);
        if (res.ok) {
          const json = await res.json();
          if (json.success && json.data) {
            return json.data;
          }
        }
      } catch (e) {
        console.warn('Node 서버 불러오기 실패:', e);
      }
    }

    if (this.firestoreDb) {
      try {
        const doc = await this.firestoreDb.collection('timetables').doc('gw_online_school').get();
        if (doc.exists) {
          return doc.data();
        }
      } catch (e) {
        console.warn('Firestore 불러오기 실패:', e);
      }
    }

    // C. 정적 호스팅(GitHub Pages) 또는 오프라인 환경인 경우: data/timetable-db.json 파일 불러오기
    try {
      const res = await fetch('data/timetable-db.json?t=' + Date.now());
      if (res.ok) {
        const json = await res.json();
        if (json && (json.yearsData || json.periods || json.students)) {
          return json;
        }
      }
    } catch (e) {
      // 정적 파일 없음 또는 네트워크 오류 (정상 동작)
    }

    return null;
  }

  // UI 뱃지 업데이트
  updateUiBadge(mode) {
    const badge = document.getElementById('backend-status-badge');
    if (!badge) return;

    if (mode === 'server') {
      badge.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800 border border-emerald-300 cursor-pointer';
      badge.innerHTML = `<span class="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span> <span>서버 DB 연결됨</span>`;
      badge.title = `Node.js 백엔드 서버 (${this.serverUrl})에 실시간 보존 중`;
    } else if (mode === 'firebase') {
      badge.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold bg-sky-100 text-sky-800 border border-sky-300 cursor-pointer';
      badge.innerHTML = `<span class="w-2 h-2 rounded-full bg-sky-500 animate-pulse"></span> <span>클라우드 DB 연결됨</span>`;
      badge.title = 'Firebase Cloud Firestore에 실시간 동기화 및 보존 중';
    } else {
      badge.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium bg-slate-100 text-slate-600 border border-slate-300 cursor-pointer';
      badge.innerHTML = `<span class="w-2 h-2 rounded-full bg-amber-400"></span> <span>로컬 저장 모드</span>`;
      badge.title = '클릭하여 백엔드 서버 또는 Firebase 클라우드 DB 연결 설정';
    }
  }

  updateStatusText(text) {
    const el = document.getElementById('backend-status-msg');
    if (el) el.textContent = text;
  }

  // Firebase 설정 저장
  saveFirebaseConfig(configStr) {
    try {
      let configObj;
      if (typeof configStr === 'string') {
        // 객체 문자열 파싱 (JSON 또는 JS 객체)
        const clean = configStr.trim()
          .replace(/^[^{]*{/, '{')
          .replace(/}[^}]*$/, '}');
        configObj = JSON.parse(clean);
      } else {
        configObj = configStr;
      }

      if (!configObj.apiKey || !configObj.projectId) {
        throw new Error('apiKey와 projectId가 필요합니다.');
      }

      localStorage.setItem(this.firebaseConfigKey, JSON.stringify(configObj));
      this.checkFirebaseConfig();
      return true;
    } catch (e) {
      alert('Firebase 설정 파싱 실패: ' + e.message);
      return false;
    }
  }

  clearFirebaseConfig() {
    localStorage.removeItem(this.firebaseConfigKey);
    this.firestoreDb = null;
    this.firebaseApp = null;
    this.detectNodeServer();
  }
}

window.cloudSync = new CloudSyncService();
