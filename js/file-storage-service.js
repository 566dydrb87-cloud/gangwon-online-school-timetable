/**
 * 강원온라인학교 시간표 시스템 - 통합 파일 저장 서비스 (FileStorageService)
 * - File System Access API 지원 브라우저(Chrome, Edge 등)에서 저장 폴더 영구 지정
 * - IndexedDB를 활용하여 선택된 폴더 핸들(DirectoryHandle)을 영구 보관 및 자동 기억
 * - JSON 백업 파일 및 모든 엑셀 다운로드 파일이 사용자가 지정한 폴더로 자동 저장
 * - 덮어쓰기 방지 기능 (동일 파일명 존재 시 (1), (2) 자동 증분)
 * - 미지원 환경 또는 폴더 미지정 시 브라우저 기본 다운로드(Blob a.click)로 자동 폴백
 */

const FileStorageService = {
  dbName: 'gw_timetable_storage_v1',
  storeName: 'settings',
  keyName: 'saveDirectoryHandle',
  cachedDirectoryHandle: null,

  /**
   * IndexedDB 열기
   */
  async openDB() {
    return new Promise((resolve, reject) => {
      if (!('indexedDB' in window)) {
        return resolve(null);
      }
      const req = indexedDB.open(this.dbName, 1);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(this.storeName)) {
          db.createObjectStore(this.storeName);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  },

  /**
   * 저장된 폴더 핸들 가져오기
   */
  async getStoredDirectoryHandle() {
    if (this.cachedDirectoryHandle) return this.cachedDirectoryHandle;
    try {
      const db = await this.openDB();
      if (!db) return null;
      return new Promise((resolve) => {
        const tx = db.transaction(this.storeName, 'readonly');
        const store = tx.objectStore(this.storeName);
        const req = store.get(this.keyName);
        req.onsuccess = () => {
          const handle = req.result;
          if (handle) {
            this.cachedDirectoryHandle = handle;
            resolve(handle);
          } else {
            resolve(null);
          }
        };
        req.onerror = () => resolve(null);
      });
    } catch (e) {
      console.warn('저장된 폴더 핸들 조회 실패:', e);
      return null;
    }
  },

  /**
   * 선택된 폴더 핸들 영구 보관
   */
  async storeDirectoryHandle(handle) {
    try {
      this.cachedDirectoryHandle = handle;
      const db = await this.openDB();
      if (!db) return false;
      return new Promise((resolve, reject) => {
        const tx = db.transaction(this.storeName, 'readwrite');
        const store = tx.objectStore(this.storeName);
        store.put(handle, this.keyName);
        tx.oncomplete = () => {
          localStorage.setItem('gw_saved_dir_name', handle.name);
          resolve(true);
        };
        tx.onerror = () => reject(tx.error);
      });
    } catch (e) {
      console.warn('폴더 핸들 IndexedDB 저장 실패:', e);
      return false;
    }
  },

  /**
   * 저장 위치 초기화 (기본 다운로드로 복원)
   */
  async clearStoredDirectoryHandle() {
    this.cachedDirectoryHandle = null;
    localStorage.removeItem('gw_saved_dir_name');
    try {
      const db = await this.openDB();
      if (!db) return;
      const tx = db.transaction(this.storeName, 'readwrite');
      tx.objectStore(this.storeName).delete(this.keyName);
    } catch (e) {}
  },

  /**
   * 현재 기억된 저장 폴더 이름 반환
   */
  getSavedDirectoryName() {
    return localStorage.getItem('gw_saved_dir_name') || null;
  },

  /**
   * 사용자에게 폴더 선택 다이얼로그 표시
   */
  async pickDirectory() {
    if (!('showDirectoryPicker' in window)) {
      throw new Error('현재 사용 중인 브라우저는 폴더 직접 지정(File System Access API)을 지원하지 않습니다. Chrome 또는 Edge 브라우저를 사용해주세요.');
    }
    const handle = await window.showDirectoryPicker({ mode: 'readwrite' });
    await this.storeDirectoryHandle(handle);
    return handle.name;
  },

  /**
   * 폴더 쓰기 권한 확인 및 재요청
   */
  async verifyPermission(fileHandle, readWrite = true) {
    if (!fileHandle) return false;
    const options = {};
    if (readWrite) options.mode = 'readwrite';
    try {
      if ((await fileHandle.queryPermission(options)) === 'granted') {
        return true;
      }
      if ((await fileHandle.requestPermission(options)) === 'granted') {
        return true;
      }
    } catch (e) {
      console.warn('권한 확인 예외:', e);
    }
    return false;
  },

  /**
   * 덮어쓰기 방지: 중복 파일명 체크 후 고유한 파일명 생성
   */
  async getUniqueFileName(dirHandle, fileName) {
    const dotIndex = fileName.lastIndexOf('.');
    const base = dotIndex !== -1 ? fileName.slice(0, dotIndex) : fileName;
    const ext = dotIndex !== -1 ? fileName.slice(dotIndex) : '';

    let candidate = fileName;
    let counter = 1;

    while (true) {
      try {
        await dirHandle.getFileHandle(candidate);
        // 파일이 존재하므로 카운터 추가 후 재확인
        candidate = `${base} (${counter})${ext}`;
        counter++;
      } catch (e) {
        // NotFoundError -> 사용 가능한 고유 파일명 확정!
        return candidate;
      }
    }
  },

  /**
   * 파일 저장 통합 메서드 (JSON, 엑셀 등 모든 파일 대상)
   * @param {ArrayBuffer|Blob|string} content 저장할 데이터 내용
   * @param {string} fileName 파일명 (확장자 포함)
   * @param {string} mimeType MIME 타입
   * @returns {Promise<{ savedToFolder: boolean, folderName: string|null, fileName: string }>}
   */
  async saveFile(content, fileName, mimeType = 'application/octet-stream') {
    let savedToFolder = false;
    let targetFileName = fileName;
    const savedDirName = this.getSavedDirectoryName();

    // 1. 저장 폴더 핸들이 등록되어 있는 경우 직접 파일 시스템에 저장 시도
    if ('showDirectoryPicker' in window && savedDirName) {
      try {
        const dirHandle = await this.getStoredDirectoryHandle();
        if (dirHandle) {
          const hasPerm = await this.verifyPermission(dirHandle, true);
          if (hasPerm) {
            // 덮어쓰기 방지 고유 파일명 생성
            targetFileName = await this.getUniqueFileName(dirHandle, fileName);
            const fileHandle = await dirHandle.getFileHandle(targetFileName, { create: true });
            const writable = await fileHandle.createWritable();
            await writable.write(content);
            await writable.close();
            savedToFolder = true;
          }
        }
      } catch (err) {
        console.warn('폴더 직접 저장 실패 -> 브라우저 기본 다운로드로 자동 전환:', err);
      }
    }

    // 2. 폴더 저장이 되지 않은 경우 기존 브라우저 표준 다운로드(Blob a.click)로 안전하게 다운로드
    if (!savedToFolder) {
      const blob = content instanceof Blob ? content : new Blob([content], { type: mimeType });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = targetFileName;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
      }, 100);
    }

    return {
      savedToFolder,
      folderName: savedDirName,
      fileName: targetFileName
    };
  }
};

window.FileStorageService = FileStorageService;
