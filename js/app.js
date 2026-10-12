/**
 * Gangwon Online School Timetable Application Main Logic
 * - 학생 편성 엑셀 업로드 시 [학번, 성명]만 추출
 * - 과목 입력 및 목록화 후 순차적 배정 (다과목 수강 지원)
 * - 3단 상태 관리: 🟢 확정 (파스텔) / 🔺 조정 (파스텔) / ✖️ 폐강 (회색 + 취소선)
 * - 원클릭 상태 전환 (테이블 & 시간표 카드 양방향)
 * - 백엔드 영구 보존 및 클라우드 동기화
 */

class TimetableApp {
  constructor() {
    this.storageKey = 'gw_online_school_timetable_data_v3';
    this.periods = [];
    this.days = [];
    this.students = [];
    this.offeredSubjects = [];
    this.subjectColorMap = {};
    this.selectedStudentIds = new Set();
    this.draggedStudentId = null;
    this.activeFilter = {
      grade: 'all',
      status: 'all',
      subject: 'all',
      search: '',
    };
    this.showWeekends = true;

    this.init();
  }

  async init() {
    this.loadState();
    this.sortPeriods();
    this.syncOfferedSubjects();
    this.buildSubjectColorMap();
    this.bindEvents();
    this.render();

    // 백엔드/클라우드 최신 데이터 확인
    setTimeout(async () => {
      if (window.cloudSync) {
        const remoteData = await window.cloudSync.loadFromBackend();
        if (remoteData && remoteData.lastUpdated) {
          const localSaved = localStorage.getItem(this.storageKey);
          const localData = localSaved ? JSON.parse(localSaved) : null;
          if (!localData || !localData.lastUpdated || new Date(remoteData.lastUpdated) > new Date(localData.lastUpdated)) {
            this.applyExternalData(remoteData, '백엔드 서버에서 최신 시간표 데이터를 불러왔습니다.');
          }
        }
      }
    }, 500);
  }

  applyExternalData(data, msg = '') {
    if (!data) return;
    if (data.periods) this.periods = data.periods;
    if (data.days) this.days = data.days;
    if (data.students) this.students = data.students;
    if (data.offeredSubjects) this.offeredSubjects = data.offeredSubjects;
    if (data.subjectColorMap) this.subjectColorMap = data.subjectColorMap;
    if (data.showWeekends !== undefined) this.showWeekends = data.showWeekends;

    this.sortPeriods();
    this.buildSubjectColorMap();
    localStorage.setItem(this.storageKey, JSON.stringify(data));
    this.render();
    if (msg) this.showToast(msg);
  }

  loadState() {
    try {
      const saved = localStorage.getItem(this.storageKey);
      if (saved) {
        const data = JSON.parse(saved);
        this.periods = data.periods || JSON.parse(JSON.stringify(DEFAULT_PERIODS));
        this.days = data.days || JSON.parse(JSON.stringify(DEFAULT_DAYS));
        this.students = data.students || [];
        this.offeredSubjects = data.offeredSubjects || [...DEFAULT_OFFERED_SUBJECTS];
        this.subjectColorMap = data.subjectColorMap || {};
        this.showWeekends = data.showWeekends !== undefined ? data.showWeekends : true;
        return;
      }
    } catch (e) {
      console.error('Failed to load state from localStorage:', e);
    }

    this.resetToDefaults(false);
  }

  saveState() {
    try {
      this.sortPeriods();
      const data = {
        periods: this.periods,
        days: this.days,
        students: this.students,
        offeredSubjects: this.offeredSubjects,
        subjectColorMap: this.subjectColorMap,
        showWeekends: this.showWeekends,
        lastUpdated: new Date().toISOString()
      };
      localStorage.setItem(this.storageKey, JSON.stringify(data));
      this.updateStats();

      // 백엔드/클라우드로 비동기 자동 저장
      if (window.cloudSync) {
        window.cloudSync.saveToBackend(data);
      }
    } catch (e) {
      console.error('Failed to save state to localStorage:', e);
    }
  }

  resetToDefaults(renderAfter = true) {
    this.periods = JSON.parse(JSON.stringify(DEFAULT_PERIODS));
    this.days = JSON.parse(JSON.stringify(DEFAULT_DAYS));
    this.students = [];
    this.offeredSubjects = [...DEFAULT_OFFERED_SUBJECTS];
    this.subjectColorMap = {};
    this.selectedStudentIds.clear();
    this.sortPeriods();
    this.buildSubjectColorMap();
    this.saveState();
    if (renderAfter) {
      this.render();
      this.showToast('데이터가 초기화되었습니다.');
    }
  }

  // 1순위: 교시 번호, 2순위: 시작 시간
  sortPeriods() {
    this.periods.sort((a, b) => this.comparePeriods(a, b));
  }

  comparePeriods(a, b) {
    const getMinutes = (timeStr) => {
      if (!timeStr) return 0;
      const [h, m] = String(timeStr).split(':').map(Number);
      return (h || 0) * 60 + (m || 0);
    };

    const pA = a.period !== null && a.period !== undefined ? Number(a.period) : null;
    const pB = b.period !== null && b.period !== undefined ? Number(b.period) : null;

    if (pA !== null && pB !== null) {
      if (pA !== pB) return pA - pB;
      return getMinutes(a.startTime) - getMinutes(b.startTime);
    }
    return getMinutes(a.startTime) - getMinutes(b.startTime);
  }

  syncOfferedSubjects() {
    this.students.forEach(s => {
      if (s.subject && s.subject !== '미지정' && !this.offeredSubjects.includes(s.subject)) {
        this.offeredSubjects.push(s.subject);
      }
    });
  }

  buildSubjectColorMap() {
    let colorIdx = Object.keys(this.subjectColorMap).length;
    this.offeredSubjects.forEach(sub => {
      if (!this.subjectColorMap[sub]) {
        this.subjectColorMap[sub] = PASTEL_PALETTE[colorIdx % PASTEL_PALETTE.length];
        colorIdx++;
      }
    });
  }

  getSubjectColor(subject) {
    if (!subject || subject === '미지정') {
      return { bg: '#F1F5F9', text: '#475569', border: '#CBD5E1', excelBg: 'F1F5F9', excelText: '475569' };
    }
    if (!this.subjectColorMap[subject]) {
      const keys = Object.keys(this.subjectColorMap);
      const color = PASTEL_PALETTE[keys.length % PASTEL_PALETTE.length];
      this.subjectColorMap[subject] = color;
      this.saveState();
    }
    return this.subjectColorMap[subject];
  }

  // ==========================================
  // 3단 상태 전환: 🟢 확정 -> 🔺 조정 -> ✖️ 폐강 -> 🟢 확정
  // ==========================================
  toggleStudentStatus(studentId) {
    const student = this.students.find(s => s.id === studentId);
    if (!student || !student.subject) return;

    if (!student.status || student.status === 'confirmed') {
      student.status = 'adjust';
    } else if (student.status === 'adjust') {
      student.status = 'canceled';
    } else {
      student.status = 'confirmed';
    }

    this.saveState();
    this.render();

    const statusLabel = ExcelService.getStatusText(student.status);
    this.showToast(`'${student.subject}(${ExcelService.getStudentCodeName(student)})' &rarr; ${statusLabel}으로 변경`);
  }

  // 이벤트 바인딩
  bindEvents() {
    // 검색 및 필터
    document.getElementById('search-input')?.addEventListener('input', (e) => {
      this.activeFilter.search = e.target.value.trim().toLowerCase();
      this.renderStudentList();
    });

    document.getElementById('filter-grade')?.addEventListener('change', (e) => {
      this.activeFilter.grade = e.target.value;
      this.renderStudentList();
    });

    document.getElementById('filter-status')?.addEventListener('change', (e) => {
      this.activeFilter.status = e.target.value;
      this.renderStudentList();
    });

    document.getElementById('filter-subject')?.addEventListener('change', (e) => {
      this.activeFilter.subject = e.target.value;
      this.renderStudentList();
    });

    // 전체 선택 체크박스
    document.getElementById('select-all-students')?.addEventListener('change', (e) => {
      const filtered = this.getFilteredStudents();
      if (e.target.checked) {
        filtered.forEach(s => this.selectedStudentIds.add(s.id));
      } else {
        filtered.forEach(s => this.selectedStudentIds.delete(s.id));
      }
      this.renderStudentList();
      this.updateSelectionBar();
    });

    // 새 과목 직접 추가 버튼 (사이드바)
    document.getElementById('btn-add-subject-direct')?.addEventListener('click', () => {
      this.handleAddSubjectFromSidebar();
    });

    document.getElementById('input-new-subject-direct')?.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        this.handleAddSubjectFromSidebar();
      }
    });

    // 일괄 배정 버튼
    document.getElementById('btn-batch-assign')?.addEventListener('click', () => {
      this.handleBatchAssign();
    });

    // 일괄 배정 해제 버튼
    document.getElementById('btn-batch-unassign')?.addEventListener('click', () => {
      this.handleBatchUnassign();
    });

    // 선택 학생 삭제 버튼
    document.getElementById('btn-batch-delete')?.addEventListener('click', () => {
      this.handleBatchDelete();
    });

    // 교시 추가 버튼
    document.getElementById('btn-add-period')?.addEventListener('click', () => {
      this.openAddPeriodModal();
    });

    document.getElementById('form-add-period')?.addEventListener('submit', (e) => {
      e.preventDefault();
      this.handleSaveCustomPeriod();
    });

    document.getElementById('btn-close-period-modal')?.addEventListener('click', () => {
      this.closeAddPeriodModal();
    });

    // 과목 관리 모달
    document.getElementById('btn-manage-subjects')?.addEventListener('click', () => {
      this.openSubjectModal();
    });

    document.getElementById('btn-close-subject-modal')?.addEventListener('click', () => {
      this.closeSubjectModal();
    });

    document.getElementById('btn-add-new-subject')?.addEventListener('click', () => {
      this.handleAddNewSubject();
    });

    // 주말 토글
    document.getElementById('btn-toggle-weekends')?.addEventListener('click', () => {
      this.showWeekends = !this.showWeekends;
      this.saveState();
      this.renderTimetable();
      this.updateWeekendToggleBtn();
    });

    // 엑셀 양식 다운로드
    document.getElementById('btn-dl-period-template')?.addEventListener('click', () => {
      ExcelService.downloadPeriodTemplate();
    });

    document.getElementById('btn-dl-student-template')?.addEventListener('click', () => {
      ExcelService.downloadStudentTemplate();
    });

    // 엑셀 업로드: 일과표
    const fileInputPeriod = document.getElementById('file-upload-period');
    fileInputPeriod?.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      try {
        const newPeriods = await ExcelService.parsePeriodExcel(file);
        if (confirm(`총 ${newPeriods.length}개의 일과표 교시 데이터를 적용하시겠습니까?`)) {
          this.periods = newPeriods;
          this.sortPeriods();
          this.saveState();
          this.render();
          this.showToast('일과표가 성공적으로 업데이트되었습니다.');
        }
      } catch (err) {
        alert('일과표 파일 처리 실패: ' + err.message);
      }
      e.target.value = '';
    });

    // 엑셀 업로드: 편성 학생 (학번, 성명만 추출)
    const fileInputStudent = document.getElementById('file-upload-student');
    fileInputStudent?.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      try {
        const newStudents = await ExcelService.parseStudentExcel(file);
        const action = confirm(
          `총 ${newStudents.length}명의 학생(학번, 성명)을 불러왔습니다!\n\n` +
          `[확인]: 기존 목록을 지우고 새로 불러온 학생들로 전체 교체\n` +
          `[취소]: 기존 목록 뒤에 추가`
        );

        if (action) {
          this.students = newStudents;
        } else {
          this.students = [...this.students, ...newStudents];
        }

        this.saveState();
        this.render();
        this.showToast(`${newStudents.length}명의 학생 명단이 등록되었습니다. 과목을 선택하고 배정해주세요.`);
      } catch (err) {
        alert('학생 명단 파일 처리 실패: ' + err.message);
      }
      e.target.value = '';
    });

    // 엑셀 결과 다운로드
    document.getElementById('btn-export-timetable')?.addEventListener('click', () => {
      const activeDays = this.showWeekends ? this.days : this.days.filter(d => !d.isWeekend);
      ExcelService.exportTimetableExcel(this.periods, activeDays, this.students, this.subjectColorMap);
      this.showToast('학급 시간표 엑셀 파일이 다운로드됩니다.');
    });

    document.getElementById('btn-export-data')?.addEventListener('click', () => {
      const activeDays = this.showWeekends ? this.days : this.days.filter(d => !d.isWeekend);
      ExcelService.exportDataExcel(this.students, this.periods, activeDays);
      this.showToast('원천 데이터 엑셀 파일이 다운로드됩니다.');
    });

    // 초기화
    document.getElementById('btn-reset-data')?.addEventListener('click', () => {
      if (confirm('모든 학생 명단과 시간표 배정을 초기화하시겠습니까?')) {
        this.resetToDefaults(true);
      }
    });

    // 개별 학생 등록 모달
    document.getElementById('btn-open-add-student-modal')?.addEventListener('click', () => {
      this.openStudentModal();
    });

    document.getElementById('form-add-student')?.addEventListener('submit', (e) => {
      e.preventDefault();
      this.handleSaveStudent();
    });

    document.getElementById('btn-close-modal')?.addEventListener('click', () => {
      this.closeStudentModal();
    });

    // 인쇄
    document.getElementById('btn-print')?.addEventListener('click', () => {
      window.print();
    });
  }

  // 사이드바에서 새 과목 입력 즉시 목록화
  handleAddSubjectFromSidebar() {
    const input = document.getElementById('input-new-subject-direct');
    const val = input.value.trim();
    if (!val) return;

    if (!this.offeredSubjects.includes(val)) {
      this.offeredSubjects.push(val);
      this.buildSubjectColorMap();
      this.saveState();
      this.populateDropdowns();
      this.showToast(`과목 '${val}'이(가) 등록되었습니다.`);
    }

    // 드롭다운에서 방금 추가한 과목 자동 선택
    const select = document.getElementById('assign-subject-select');
    if (select) select.value = val;
    input.value = '';
  }

  getFilteredStudents() {
    return this.students.filter(student => {
      if (this.activeFilter.grade !== 'all' && String(student.grade) !== this.activeFilter.grade) {
        return false;
      }
      if (this.activeFilter.subject !== 'all' && student.subject !== this.activeFilter.subject) {
        return false;
      }
      const isAssigned = student.day && (student.periodId || student.period);
      if (this.activeFilter.status === 'assigned' && !isAssigned) {
        return false;
      }
      if (this.activeFilter.status === 'unassigned' && isAssigned) {
        return false;
      }
      if (this.activeFilter.status === 'confirmed' && (!isAssigned || student.status !== 'confirmed')) {
        return false;
      }
      if (this.activeFilter.status === 'adjust' && (!isAssigned || student.status !== 'adjust')) {
        return false;
      }
      if (this.activeFilter.status === 'canceled' && (!isAssigned || student.status !== 'canceled')) {
        return false;
      }

      if (this.activeFilter.search) {
        const query = this.activeFilter.search;
        const code = ExcelService.getStudentCode(student);
        const codeName = ExcelService.getStudentCodeName(student).toLowerCase();
        const matchesName = student.name?.toLowerCase().includes(query);
        const matchesCode = code.includes(query);
        const matchesSubject = student.subject?.toLowerCase().includes(query);
        if (!matchesName && !matchesCode && !matchesCodeName && !matchesSubject) return false;
      }
      return true;
    });
  }

  render() {
    this.updateStats();
    this.populateDropdowns();
    this.renderTimetable();
    this.renderStudentList();
    this.updateWeekendToggleBtn();
    this.updateSelectionBar();
  }

  updateStats() {
    const total = this.students.length;
    const assigned = this.students.filter(s => s.day && (s.periodId || s.period)).length;
    const unassigned = total - assigned;
    const regularPeriods = this.periods.filter(p => !p.isBreak).length;

    const elTotal = document.getElementById('stat-total-students');
    const elAssigned = document.getElementById('stat-assigned-students');
    const elUnassigned = document.getElementById('stat-unassigned-students');
    const elPeriods = document.getElementById('stat-period-count');

    if (elTotal) elTotal.textContent = total;
    if (elAssigned) elAssigned.textContent = assigned;
    if (elUnassigned) elUnassigned.textContent = unassigned;
    if (elPeriods) elPeriods.textContent = regularPeriods;
  }

  populateDropdowns() {
    // 1. 배정 과목 선택 드롭다운
    const assignSubjectSelect = document.getElementById('assign-subject-select');
    if (assignSubjectSelect) {
      const cur = assignSubjectSelect.value;
      let optionsHtml = `<option value="">-- [과목 선택] --</option>`;
      this.offeredSubjects.forEach(sub => {
        optionsHtml += `<option value="${sub}" ${cur === sub ? 'selected' : ''}>${sub}</option>`;
      });
      assignSubjectSelect.innerHTML = optionsHtml;
    }

    // 2. 배정 요일 드롭다운
    const assignDaySelect = document.getElementById('assign-day-select');
    if (assignDaySelect) {
      const cur = assignDaySelect.value;
      assignDaySelect.innerHTML = this.days.map(d => 
        `<option value="${d.id}" ${cur === d.id ? 'selected' : ''}>${d.fullName || d.name + '요일'}${d.isWeekend ? ' (주말)' : ''}</option>`
      ).join('');
    }

    // 3. 배정 교시 드롭다운
    const assignPeriodSelect = document.getElementById('assign-period-select');
    if (assignPeriodSelect) {
      const cur = assignPeriodSelect.value;
      const regularPeriods = this.periods.filter(p => !p.isBreak);
      assignPeriodSelect.innerHTML = regularPeriods.map(p => 
        `<option value="${p.id}" ${cur === p.id ? 'selected' : ''}>${p.name} (${p.startTime}~${p.endTime})</option>`
      ).join('');
    }

    // 4. 과목 필터 드롭다운
    const filterSubjectSelect = document.getElementById('filter-subject');
    if (filterSubjectSelect) {
      const cur = this.activeFilter.subject;
      let filterHtml = `<option value="all">전체 과목</option>`;
      this.offeredSubjects.forEach(sub => {
        filterHtml += `<option value="${sub}" ${cur === sub ? 'selected' : ''}>${sub}</option>`;
      });
      filterSubjectSelect.innerHTML = filterHtml;
    }
  }

  updateWeekendToggleBtn() {
    const btn = document.getElementById('btn-toggle-weekends');
    if (!btn) return;
    if (this.showWeekends) {
      btn.innerHTML = `<i data-lucide="calendar" class="w-4 h-4 inline-block mr-1"></i> 주말 숨기기(월~금)`;
      btn.classList.add('bg-blue-50', 'text-blue-700', 'border-blue-200');
    } else {
      btn.innerHTML = `<i data-lucide="calendar-plus" class="w-4 h-4 inline-block mr-1"></i> 주말 포함(월~일)`;
      btn.classList.remove('bg-blue-50', 'text-blue-700', 'border-blue-200');
    }
    if (window.lucide) window.lucide.createIcons();
  }

  // ==========================================
  // 시간표 그리드 렌더링 & 드래그 앤 드롭 & 카드 상태 표기
  // ==========================================
  renderTimetable() {
    const thead = document.getElementById('timetable-thead');
    const tbody = document.getElementById('timetable-tbody');
    if (!thead || !tbody) return;

    const activeDays = this.showWeekends ? this.days : this.days.filter(d => !d.isWeekend);

    let headerHtml = `
      <tr>
        <th class="py-3 px-3 w-36 bg-slate-100 text-slate-700 font-semibold border border-slate-200 text-center text-xs tracking-wider sticky left-0 z-20">
          교시 / 시간
        </th>
    `;
    activeDays.forEach(day => {
      const isW = day.isWeekend;
      headerHtml += `
        <th class="py-3 px-3 min-w-[140px] font-semibold border border-slate-200 text-center text-sm ${isW ? 'bg-rose-50 text-rose-700' : 'bg-blue-50 text-blue-900'}">
          <div class="flex items-center justify-center gap-1.5">
            <span>${day.fullName || day.name + '요일'}</span>
            ${isW ? '<span class="text-[10px] px-1.5 py-0.5 rounded bg-rose-200/70 text-rose-800 font-normal">주말</span>' : ''}
          </div>
        </th>
      `;
    });
    headerHtml += `</tr>`;
    thead.innerHTML = headerHtml;

    let bodyHtml = '';

    this.periods.forEach((period) => {
      if (period.isBreak) {
        bodyHtml += `
          <tr class="bg-slate-50/80">
            <td class="py-2 px-3 border border-slate-200 text-center bg-slate-100/70 sticky left-0 z-10">
              <div class="flex items-center justify-center gap-1 font-semibold text-slate-600 text-xs">
                <span>${period.name}</span>
                <button type="button" onclick="window.app.handleDeletePeriod('${period.id}')" title="삭제" class="text-slate-400 hover:text-red-500 p-0.5">
                  <i data-lucide="x" class="w-3 h-3"></i>
                </button>
              </div>
              <div class="flex items-center justify-center gap-0.5 mt-0.5">
                <input type="text" class="time-input-inline" value="${period.startTime}" 
                  onchange="window.app.updatePeriodTime('${period.id}', 'startTime', this.value)" title="시작 시간">
                <span class="text-slate-400 text-xs">~</span>
                <input type="text" class="time-input-inline" value="${period.endTime}" 
                  onchange="window.app.updatePeriodTime('${period.id}', 'endTime', this.value)" title="종료 시간">
              </div>
            </td>
            <td colspan="${activeDays.length}" class="py-2.5 px-4 border border-slate-200 text-center text-slate-500 text-xs font-medium">
              <div class="inline-flex items-center gap-2 px-3 py-1 bg-slate-100 rounded-full border border-slate-200/60">
                <i data-lucide="utensils" class="w-3.5 h-3.5 text-amber-600"></i>
                <span>${period.name} (${period.startTime} ~ ${period.endTime})</span>
              </div>
            </td>
          </tr>
        `;
      } else {
        bodyHtml += `
          <tr class="hover:bg-slate-50/30 transition-colors">
            <td class="py-2.5 px-2 border border-slate-200 text-center bg-slate-50/80 sticky left-0 z-10 align-middle">
              <div class="flex items-center justify-center gap-1">
                <span class="font-bold text-slate-800 text-xs">${period.name}</span>
                <button type="button" onclick="window.app.handleDeletePeriod('${period.id}')" title="이 교시 삭제" class="text-slate-300 hover:text-red-500 p-0.5 rounded transition">
                  <i data-lucide="x" class="w-3 h-3"></i>
                </button>
              </div>
              <div class="flex items-center justify-center gap-0.5 mt-1">
                <input type="text" class="time-input-inline" value="${period.startTime}" 
                  onchange="window.app.updatePeriodTime('${period.id}', 'startTime', this.value)" title="시작 시간">
                <span class="text-slate-400 text-xs">~</span>
                <input type="text" class="time-input-inline" value="${period.endTime}" 
                  onchange="window.app.updatePeriodTime('${period.id}', 'endTime', this.value)" title="종료 시간">
              </div>
            </td>
        `;

        activeDays.forEach(day => {
          const assigned = this.students.filter(s => {
            if (s.day !== day.id) return false;
            if (s.periodId) return s.periodId === period.id;
            return s.period === period.period;
          });

          bodyHtml += `
            <td class="timetable-cell p-2 border border-slate-200 align-top transition-colors"
                data-day="${day.id}" 
                data-period-id="${period.id}"
                ondragover="window.app.handleDragOver(event)"
                ondragleave="window.app.handleDragLeave(event)"
                ondrop="window.app.handleDrop(event, '${day.id}', '${period.id}')">
              <div class="flex flex-col gap-1.5 min-h-[60px]">
          `;

          if (assigned.length > 0) {
            assigned.forEach(student => {
              const isCanceled = student.status === 'canceled';
              const color = this.getSubjectColor(student.subject);
              const symbol = ExcelService.getStatusSymbol(student.status);
              const codeName = ExcelService.getStudentCodeName(student);

              // 폐강일 경우: 회색 배경 + 취소선 스타일
              const bgStyle = isCanceled ? '#E2E8F0' : color.bg;
              const textStyle = isCanceled ? '#64748B' : color.text;
              const borderStyle = isCanceled ? '#CBD5E1' : color.border;
              const strikeClass = isCanceled ? 'line-through opacity-75' : '';

              bodyHtml += `
                <div class="subject-card badge-pill shadow-xs group ${strikeClass}" 
                     draggable="true"
                     data-student-id="${student.id}"
                     ondragstart="window.app.handleDragStart(event, '${student.id}')"
                     ondragend="window.app.handleDragEnd(event)"
                     style="background-color: ${bgStyle}; color: ${textStyle}; border-color: ${borderStyle};"
                     title="학번성명: ${codeName} / 과목: ${student.subject} / 상태: ${ExcelService.getStatusText(student.status)} (클릭하여 상태 변경)">
                  
                  <!-- 상태 이모티콘 (원클릭 순환 변경) -->
                  <span onclick="event.stopPropagation(); window.app.toggleStudentStatus('${student.id}')" 
                        class="cursor-pointer mr-1 text-[11px] hover:scale-125 transition-transform" 
                        title="상태 클릭 변경 (🟢확정 / 🔺조정 / ✖️폐강)">
                    ${symbol}
                  </span>

                  <span class="truncate max-w-[140px] text-xs font-semibold">
                    ${student.subject}(${codeName})
                  </span>

                  <button type="button" 
                          onclick="event.stopPropagation(); window.app.unassignStudent('${student.id}')" 
                          class="ml-1.5 opacity-60 hover:opacity-100 hover:text-red-700 transition" 
                          title="배정 해제">
                    &times;
                  </button>
                </div>
              `;
            });
          }

          bodyHtml += `
              </div>
            </td>
          `;
        });

        bodyHtml += `</tr>`;
      }
    });

    tbody.innerHTML = bodyHtml;
    if (window.lucide) window.lucide.createIcons();
  }

  updatePeriodTime(periodId, field, value) {
    const p = this.periods.find(item => item.id === periodId);
    if (!p) return;
    p[field] = value.trim();
    this.sortPeriods();
    this.saveState();
    this.render();
  }

  // ==========================================
  // 사용자 지정 교시 추가 모달
  // ==========================================
  openAddPeriodModal() {
    const modal = document.getElementById('period-modal');
    if (!modal) return;
    
    const regularPeriods = this.periods.filter(p => !p.isBreak);
    const lastPeriod = regularPeriods[regularPeriods.length - 1];
    const nextPeriodNum = lastPeriod ? (lastPeriod.period || 8) + 1 : 1;

    let defStart = '09:00';
    let defEnd = '09:50';

    if (lastPeriod && lastPeriod.endTime) {
      const [h, m] = lastPeriod.endTime.split(':').map(Number);
      if (!isNaN(h) && !isNaN(m)) {
        const startMins = h * 60 + m + 10;
        const endMins = startMins + 50;
        defStart = `${String(Math.floor(startMins / 60)).padStart(2, '0')}:${String(startMins % 60).padStart(2, '0')}`;
        defEnd = `${String(Math.floor(endMins / 60)).padStart(2, '0')}:${String(endMins % 60).padStart(2, '0')}`;
      }
    }

    document.getElementById('modal-period-num').value = nextPeriodNum;
    document.getElementById('modal-period-name').value = `${nextPeriodNum}교시`;
    document.getElementById('modal-period-start').value = defStart;
    document.getElementById('modal-period-end').value = defEnd;
    document.getElementById('modal-period-type').value = 'class';

    modal.classList.remove('hidden');
    modal.classList.add('flex');
  }

  closeAddPeriodModal() {
    const modal = document.getElementById('period-modal');
    if (!modal) return;
    modal.classList.add('hidden');
    modal.classList.remove('flex');
  }

  handleSaveCustomPeriod() {
    const type = document.getElementById('modal-period-type').value;
    const isBreak = type === 'break';
    const periodNum = isBreak ? null : parseInt(document.getElementById('modal-period-num').value, 10);
    const name = document.getElementById('modal-period-name').value.trim();
    const startTime = document.getElementById('modal-period-start').value.trim();
    const endTime = document.getElementById('modal-period-end').value.trim();

    if (!name || !startTime || !endTime) {
      alert('교시명과 시작/종료 시간을 모두 입력해주세요.');
      return;
    }

    const newPeriod = {
      id: `p_custom_${Date.now()}`,
      period: periodNum,
      name: name,
      startTime: startTime,
      endTime: endTime,
      isBreak: isBreak
    };

    this.periods.push(newPeriod);
    this.sortPeriods();
    this.saveState();
    this.render();
    this.closeAddPeriodModal();
    this.showToast(`'${name}' 교시가 추가되었습니다.`);
  }

  handleDeletePeriod(periodId) {
    const period = this.periods.find(p => p.id === periodId);
    if (!period) return;

    const assigned = this.students.filter(s => s.periodId === period.id || (!s.periodId && s.period === period.period));
    if (assigned.length > 0) {
      if (!confirm(`이 교시에는 ${assigned.length}명의 배정 항목이 있습니다. 삭제하면 배정이 해제됩니다. 계속하시겠습니까?`)) {
        return;
      }
      assigned.forEach(s => {
        s.day = null;
        s.period = null;
        s.periodId = null;
      });
    }

    this.periods = this.periods.filter(p => p.id !== periodId);
    this.sortPeriods();
    this.saveState();
    this.render();
    this.showToast(`${period.name}가 삭제되었습니다.`);
  }

  // ==========================================
  // 개설 과목 관리 모달
  // ==========================================
  openSubjectModal() {
    const modal = document.getElementById('subject-modal');
    if (!modal) return;
    this.renderSubjectModalList();
    modal.classList.remove('hidden');
    modal.classList.add('flex');
  }

  closeSubjectModal() {
    const modal = document.getElementById('subject-modal');
    if (!modal) return;
    modal.classList.add('hidden');
    modal.classList.remove('flex');
  }

  renderSubjectModalList() {
    const listEl = document.getElementById('subject-modal-list');
    if (!listEl) return;

    if (this.offeredSubjects.length === 0) {
      listEl.innerHTML = `<p class="text-xs text-slate-400 py-4 text-center">등록된 과목이 없습니다.</p>`;
      return;
    }

    listEl.innerHTML = this.offeredSubjects.map(sub => {
      const color = this.getSubjectColor(sub);
      const studentCount = this.students.filter(s => s.subject === sub).length;
      return `
        <div class="flex items-center justify-between p-2 rounded-lg border border-slate-200 bg-white">
          <div class="flex items-center gap-2">
            <span class="w-3.5 h-3.5 rounded-full border" style="background-color: ${color.bg}; border-color: ${color.border};"></span>
            <span class="font-bold text-xs text-slate-800">${sub}</span>
            <span class="text-[11px] text-slate-400">(${studentCount}명 배정)</span>
          </div>
          <button type="button" onclick="window.app.handleDeleteSubject('${sub}')" class="text-slate-400 hover:text-red-500 p-1 text-xs">
            삭제
          </button>
        </div>
      `;
    }).join('');
  }

  handleAddNewSubject() {
    const input = document.getElementById('modal-new-subject-input');
    const val = input.value.trim();
    if (!val) return;
    if (this.offeredSubjects.includes(val)) {
      alert('이미 등록된 과목입니다.');
      return;
    }
    this.offeredSubjects.push(val);
    this.buildSubjectColorMap();
    this.saveState();
    this.populateDropdowns();
    this.renderSubjectModalList();
    input.value = '';
    this.showToast(`과목 '${val}'이(가) 등록되었습니다.`);
  }

  handleDeleteSubject(subName) {
    if (!confirm(`'${subName}' 과목을 목록에서 삭제하시겠습니까?`)) return;
    this.offeredSubjects = this.offeredSubjects.filter(s => s !== subName);
    this.saveState();
    this.populateDropdowns();
    this.renderSubjectModalList();
    this.render();
  }

  // ==========================================
  // 드래그 앤 드롭
  // ==========================================
  handleDragStart(e, studentId) {
    this.draggedStudentId = studentId;
    e.dataTransfer.setData('text/plain', studentId);
    e.dataTransfer.effectAllowed = 'move';
    e.target.classList.add('is-dragging');
  }

  handleDragEnd(e) {
    this.draggedStudentId = null;
    e.target.classList.remove('is-dragging');
    document.querySelectorAll('.timetable-cell.drag-over').forEach(cell => {
      cell.classList.remove('drag-over');
    });
  }

  handleDragOver(e) {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    const cell = e.currentTarget;
    if (!cell.classList.contains('drag-over')) {
      cell.classList.add('drag-over');
    }
  }

  handleDragLeave(e) {
    const cell = e.currentTarget;
    cell.classList.remove('drag-over');
  }

  handleDrop(e, targetDay, targetPeriodId) {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-over');

    const studentId = this.draggedStudentId || e.dataTransfer.getData('text/plain');
    if (!studentId) return;

    const student = this.students.find(s => s.id === studentId);
    if (!student) return;

    const periodObj = this.periods.find(p => p.id === targetPeriodId);
    if (!periodObj) return;

    student.day = targetDay;
    student.periodId = periodObj.id;
    student.period = periodObj.period;

    this.saveState();
    this.renderTimetable();
    this.renderStudentList();

    const codeName = ExcelService.getStudentCodeName(student);
    this.showToast(`${student.subject}(${codeName}) &rarr; ${ExcelService.dayIdToName(targetDay)}요일 ${periodObj.name}로 이동`);
  }

  unassignStudent(studentId) {
    const student = this.students.find(s => s.id === studentId);
    if (!student) return;
    student.day = null;
    student.period = null;
    student.periodId = null;
    this.saveState();
    this.renderTimetable();
    this.renderStudentList();
    this.showToast(`${student.name} 학생의 배정이 해제되었습니다.`);
  }

  // ==========================================
  // 학생 목록 & 순차 배정 실행 (다과목 수강 지원)
  // ==========================================
  renderStudentList() {
    const tbody = document.getElementById('student-list-tbody');
    if (!tbody) return;

    const filtered = this.getFilteredStudents();

    if (filtered.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="7" class="py-10 text-center text-slate-400 text-xs">
            <i data-lucide="users" class="w-8 h-8 mx-auto mb-2 opacity-40"></i>
            <p class="font-medium text-slate-600">등록된 학생 데이터가 없습니다.</p>
            <p class="text-[11px] text-slate-400 mt-1">상단의 <strong>[엑셀 업로드]</strong>로 명단을 등록해주세요.</p>
          </td>
        </tr>
      `;
      if (window.lucide) window.lucide.createIcons();
      return;
    }

    let rowsHtml = '';
    filtered.forEach(student => {
      const isSelected = this.selectedStudentIds.has(student.id);
      const isAssigned = student.day && (student.periodId || student.period);
      const isCanceled = student.status === 'canceled';
      const studentCode = student.studentCode || ExcelService.getStudentCode(student);
      const color = this.getSubjectColor(student.subject);
      const symbol = ExcelService.getStatusSymbol(student.status);

      const periodObj = this.periods.find(p => p.id === student.periodId || (!student.periodId && p.period === student.period));
      const periodLabel = periodObj ? periodObj.name : (student.period ? `${student.period}교시` : '');

      rowsHtml += `
        <tr class="hover:bg-slate-50 border-b border-slate-100 transition-colors ${isSelected ? 'bg-blue-50/50' : ''} ${isCanceled ? 'bg-slate-100/50 text-slate-400' : ''}">
          <td class="py-2 px-2 text-center">
            <input type="checkbox" value="${student.id}" ${isSelected ? 'checked' : ''} 
                   onchange="window.app.toggleSelectStudent('${student.id}', this.checked)">
          </td>
          <!-- 학번 -->
          <td class="py-2 px-1 text-center text-xs font-mono font-bold text-slate-700 whitespace-nowrap">
            ${studentCode}
          </td>
          <!-- 성명 -->
          <td class="py-2 px-2 text-left font-semibold text-xs text-slate-800 whitespace-nowrap ${isCanceled ? 'line-through' : ''}">
            ${student.name}
          </td>
          <!-- 과목명 -->
          <td class="py-2 px-2 text-left text-xs whitespace-nowrap">
            ${student.subject ? `
              <span class="inline-block px-2 py-0.5 rounded text-[11px] font-semibold border ${isCanceled ? 'line-through' : ''}"
                    style="background-color: ${isCanceled ? '#E2E8F0' : color.bg}; color: ${isCanceled ? '#64748B' : color.text}; border-color: ${isCanceled ? '#CBD5E1' : color.border};">
                ${student.subject}
              </span>
            ` : `
              <span class="text-slate-300 text-[11px]">-</span>
            `}
          </td>
          <!-- 배정 교시 -->
          <td class="py-2 px-1 text-center text-xs whitespace-nowrap">
            ${isAssigned ? `
              <span class="inline-flex items-center gap-1 text-[11px] px-2 py-0.5 rounded-full ${isCanceled ? 'bg-slate-200 text-slate-500' : 'bg-emerald-50 text-emerald-700'} font-medium border border-slate-200">
                ${ExcelService.dayIdToName(student.day)} ${periodLabel}
              </span>
            ` : `
              <span class="text-[11px] px-1.5 py-0.5 rounded text-slate-400 bg-slate-100 font-normal">
                미배정
              </span>
            `}
          </td>
          <!-- 상태 (🟢 확정 / 🔺 조정 / ✖️ 폐강 원클릭 전환) -->
          <td class="py-2 px-1 text-center text-xs whitespace-nowrap">
            ${isAssigned ? `
              <button type="button" 
                      onclick="window.app.toggleStudentStatus('${student.id}')"
                      class="inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-[11px] font-bold border transition hover:scale-105 ${
                        student.status === 'canceled' ? 'bg-slate-200 text-slate-600 border-slate-300' :
                        student.status === 'adjust' ? 'bg-amber-100 text-amber-800 border-amber-300' :
                        'bg-emerald-100 text-emerald-800 border-emerald-300'
                      }"
                      title="클릭하여 상태 변경 (🟢확정 -> 🔺조정 -> ✖️폐강)">
                <span>${symbol}</span>
                <span class="text-[10px]">${student.status === 'canceled' ? '폐강' : student.status === 'adjust' ? '조정' : '확정'}</span>
              </button>
            ` : `
              <span class="text-slate-300 text-xs">-</span>
            `}
          </td>
          <!-- 삭제 -->
          <td class="py-2 px-1 text-center text-xs whitespace-nowrap">
            <button type="button" onclick="window.app.deleteStudent('${student.id}')" title="삭제" class="text-slate-400 hover:text-red-500 p-1 rounded transition">
              <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
            </button>
          </td>
        </tr>
      `;
    });

    tbody.innerHTML = rowsHtml;
    if (window.lucide) window.lucide.createIcons();
  }

  toggleSelectStudent(studentId, isChecked) {
    if (isChecked) {
      this.selectedStudentIds.add(studentId);
    } else {
      this.selectedStudentIds.delete(studentId);
    }
    this.updateSelectionBar();
  }

  updateSelectionBar() {
    const count = this.selectedStudentIds.size;
    const countBadge = document.getElementById('selected-count-badge');
    const assignBtn = document.getElementById('btn-batch-assign');
    const unassignBtn = document.getElementById('btn-batch-unassign');
    const deleteBtn = document.getElementById('btn-batch-delete');

    if (countBadge) countBadge.textContent = `${count}명 선택됨`;

    if (assignBtn) assignBtn.disabled = count === 0;
    if (unassignBtn) unassignBtn.disabled = count === 0;
    if (deleteBtn) deleteBtn.disabled = count === 0;

    const selectAllBox = document.getElementById('select-all-students');
    const filtered = this.getFilteredStudents();
    if (selectAllBox && filtered.length > 0) {
      selectAllBox.checked = filtered.every(s => this.selectedStudentIds.has(s.id));
    }
  }

  // 순차 배정 실행 (학생 체크 -> 과목 선택 -> 시간/요일 선택 -> 배정)
  handleBatchAssign() {
    if (this.selectedStudentIds.size === 0) {
      alert('배정할 학생을 먼저 체크 선택해주세요.');
      return;
    }

    const subjectSelect = document.getElementById('assign-subject-select');
    const daySelect = document.getElementById('assign-day-select');
    const periodSelect = document.getElementById('assign-period-select');

    const assignedSubject = subjectSelect.value.trim();
    const dayId = daySelect.value;
    const periodId = periodSelect.value;
    const periodObj = this.periods.find(p => p.id === periodId);

    if (!assignedSubject) {
      alert('배정할 과목을 먼저 선택하거나 직접 입력하여 추가해주세요.');
      subjectSelect.focus();
      return;
    }

    if (!periodObj) {
      alert('배정할 교시를 선택해주세요.');
      return;
    }

    let assignedCount = 0;

    // 선택된 학생들에 대해 다과목 수강 배정 처리
    const newEnrollments = [];

    this.students.forEach(s => {
      if (this.selectedStudentIds.has(s.id)) {
        // 아직 과목이 없는 최초 학생인 경우: 해당 학생 행에 바로 과목/시간 지정
        if (!s.subject) {
          s.subject = assignedSubject;
          s.day = dayId;
          s.periodId = periodObj.id;
          s.period = periodObj.period;
          s.status = 'confirmed';
          assignedCount++;
        } else {
          // 이미 과목이 배정되어 있는 학생인 경우: 다과목 수강을 위해 새 수강 항목으로 추가!
          const newEntry = {
            ...s,
            id: `s_multi_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`,
            subject: assignedSubject,
            day: dayId,
            periodId: periodObj.id,
            period: periodObj.period,
            status: 'confirmed'
          };
          newEnrollments.push(newEntry);
          assignedCount++;
        }
      }
    });

    if (newEnrollments.length > 0) {
      this.students.push(...newEnrollments);
    }

    this.selectedStudentIds.clear();
    this.saveState();
    this.render();
    this.showToast(`선택한 ${assignedCount}명의 학생에게 '${assignedSubject}' (${ExcelService.dayIdToName(dayId)} ${periodObj.name}) 과목이 배정되었습니다.`);
  }

  handleBatchUnassign() {
    if (this.selectedStudentIds.size === 0) return;

    let count = 0;
    this.students.forEach(s => {
      if (this.selectedStudentIds.has(s.id)) {
        s.day = null;
        s.period = null;
        s.periodId = null;
        count++;
      }
    });

    this.selectedStudentIds.clear();
    this.saveState();
    this.render();
    this.showToast(`선택한 ${count}명의 배정이 해제되었습니다.`);
  }

  handleBatchDelete() {
    if (this.selectedStudentIds.size === 0) return;
    if (!confirm(`선택한 ${this.selectedStudentIds.size}명의 항목을 삭제하시겠습니까?`)) return;

    this.students = this.students.filter(s => !this.selectedStudentIds.has(s.id));
    this.selectedStudentIds.clear();
    this.saveState();
    this.render();
    this.showToast('선택한 항목이 삭제되었습니다.');
  }

  deleteStudent(studentId) {
    const s = this.students.find(item => item.id === studentId);
    if (!s) return;
    if (confirm(`'${s.name} ${s.subject ? '(' + s.subject + ')' : ''}' 학생 항목을 삭제하시겠습니까?`)) {
      this.students = this.students.filter(item => item.id !== studentId);
      this.selectedStudentIds.delete(studentId);
      this.saveState();
      this.render();
      this.showToast(`${s.name} 학생 항목이 삭제되었습니다.`);
    }
  }

  // 개별 학생 등록 모달
  openStudentModal() {
    const modal = document.getElementById('student-modal');
    if (!modal) return;
    document.getElementById('form-add-student').reset();
    modal.classList.remove('hidden');
    modal.classList.add('flex');
  }

  closeStudentModal() {
    const modal = document.getElementById('student-modal');
    if (!modal) return;
    modal.classList.add('hidden');
    modal.classList.remove('flex');
  }

  handleSaveStudent() {
    const grade = parseInt(document.getElementById('modal-grade').value, 10) || 1;
    const classNum = parseInt(document.getElementById('modal-class').value, 10) || 1;
    const studentNum = parseInt(document.getElementById('modal-number').value, 10) || null;
    const name = document.getElementById('modal-name').value.trim();

    if (!name) {
      alert('학생 성명을 입력해주세요.');
      return;
    }

    const studentCode = ExcelService.getStudentCode({ grade, classNum, studentNum });

    const newStudent = {
      id: `s_${Date.now()}`,
      grade: grade,
      classNum: classNum,
      studentNum: studentNum,
      studentCode: studentCode,
      name: name,
      subject: '', // 최초 등록 시 과목 미배정
      day: null,
      period: null,
      periodId: null,
      status: 'confirmed'
    };

    this.students.push(newStudent);
    this.saveState();
    this.render();
    this.closeStudentModal();
    this.showToast(`'${studentCode} ${name}' 학생이 등록되었습니다.`);
  }

  showToast(message) {
    let toast = document.getElementById('app-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'app-toast';
      toast.className = 'fixed bottom-5 right-5 z-50 bg-slate-900 text-white text-xs px-4 py-2.5 rounded-lg shadow-lg flex items-center gap-2 transform transition-all duration-300 opacity-0 translate-y-3';
      document.body.appendChild(toast);
    }
    toast.innerHTML = `<i data-lucide="check-circle" class="w-4 h-4 text-emerald-400"></i> <span>${message}</span>`;
    if (window.lucide) window.lucide.createIcons();

    toast.classList.remove('opacity-0', 'translate-y-3');
    toast.classList.add('opacity-100', 'translate-y-0');

    if (this.toastTimeout) clearTimeout(this.toastTimeout);
    this.toastTimeout = setTimeout(() => {
      toast.classList.remove('opacity-100', 'translate-y-0');
      toast.classList.add('opacity-0', 'translate-y-3');
    }, 2800);
  }
}

document.addEventListener('DOMContentLoaded', () => {
  window.app = new TimetableApp();
});
