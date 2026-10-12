/**
 * Gangwon Online School Timetable Application Main Logic
 * 
 * [새로운 2단계 작업 방식]
 * 1단계 (탭 1): 학생·과목 등록 및 과목별 편제 (바구니/폴더)
 *   - 학생 명단 (학생 풀)에서 학생 카드를 마우스로 드래그하여 과목 바구니(Folder/Bucket)에 Drop하면 편제.
 *   - 과목 바구니 내에서 시간(요일/교시)을 추가 설정하면 시간표에 자동 배치 (여러 날, 여러 시간 지원).
 * 
 * 2단계 (탭 2): 종합 시간표 편성 (과목 드래그 배치 & 복사)
 *   - 편제된 과목 팔레트에서 과목 카드를 시간표 셀로 드래그하여 배치.
 *   - 배치 시 과목에 편제된 모든 학생 정보가 함께 자동 할당.
 *   - 시간표 내 카드를 드래그하여 다른 교시/요일로 이동(Move).
 *   - Ctrl 키를 누른 상태에서 드래그하면 원하는 시간에 즉시 복사(Copy)되어 복수 교시 편성.
 */

class TimetableApp {
  constructor() {
    this.storageKey = 'gw_online_school_timetable_data_v4';
    this.legacyStorageKey = 'gw_online_school_timetable_data_v3';
    this.activeYear = '2026';
    this.activeSemester = '1';
    this.academicYears = ['2026', '2027'];
    this.currentWorkflowTab = 'step1'; // 'step1' | 'step2'
    this.yearsData = {};

    // 현재 활성 학년도 및 학기의 데이터
    this.periods = [];
    this.days = [];
    this.masterStudents = []; // [{ id, grade, classNum, studentNum, studentCode, name }]
    this.courseEnrollments = {}; // { [subjectName]: [studentId, ...] }
    this.courseSchedules = {}; // { [subjectName]: [{ id, day, periodId, status }, ...] }
    this.students = []; // flat assignments array for timetable & ExcelService
    this.offeredSubjects = [];
    this.subjectColorMap = {};
    this.showWeekends = true;

    // UI 상태 관리
    this.selectedStep1StudentIds = new Set();
    this.step1Filter = {
      search: '',
      grade: 'all',
      enroll: 'all'
    };
    this.isCtrlPressed = false;
    this.dragData = null; // { type, studentId, studentIds, subjectName, fromDay, fromPeriodId }

    this.init();
  }

  async init() {
    this.loadState();
    this.sortPeriods();
    this.syncOfferedSubjects();
    this.buildSubjectColorMap();
    this.bindEvents();
    this.render();

    // 백엔드/클라우드 최신 데이터 자동 동기화
    setTimeout(async () => {
      if (window.cloudSync) {
        const remoteData = await window.cloudSync.loadFromBackend();
        if (remoteData) {
          const localSaved = localStorage.getItem(this.storageKey);
          const localData = localSaved ? JSON.parse(localSaved) : null;
          const localCount = (this.masterStudents || []).length;

          // 원격 데이터가 더 최신이거나, 로컬 학생 데이터가 비어있는 경우 즉시 적용
          const shouldApply = !localData || 
            localCount === 0 || 
            !localData.lastUpdated || 
            (remoteData.lastUpdated && new Date(remoteData.lastUpdated) > new Date(localData.lastUpdated));

          if (shouldApply) {
            this.applyExternalData(remoteData, '시간표 및 학생 데이터를 불러왔습니다.');
          }
        }
      }
    }, 200);
  }

  // ==========================================
  // 데이터 구조 생성 및 마이그레이션
  // ==========================================
  createDefaultSemesterData(basePeriods = null, baseDays = null, baseSubjects = null, baseColorMap = null) {
    return {
      periods: basePeriods ? JSON.parse(JSON.stringify(basePeriods)) : JSON.parse(JSON.stringify(DEFAULT_PERIODS)),
      days: baseDays ? JSON.parse(JSON.stringify(baseDays)) : JSON.parse(JSON.stringify(DEFAULT_DAYS)),
      masterStudents: [],
      courseEnrollments: {},
      courseSchedules: {},
      students: [],
      offeredSubjects: baseSubjects ? [...baseSubjects] : [...DEFAULT_OFFERED_SUBJECTS],
      subjectColorMap: baseColorMap ? JSON.parse(JSON.stringify(baseColorMap)) : {},
      showWeekends: true
    };
  }

  ensureYearSemesterData(year, semester = '1') {
    if (!this.yearsData[year]) {
      this.yearsData[year] = {
        '1': this.createDefaultSemesterData(),
        '2': this.createDefaultSemesterData()
      };
    }

    // 구 버전 데이터 구조 마이그레이션
    if (this.yearsData[year].periods || this.yearsData[year].students) {
      const old = this.yearsData[year];
      this.yearsData[year] = {
        '1': old,
        '2': this.createDefaultSemesterData(old.periods, old.days, old.offeredSubjects, old.subjectColorMap)
      };
    }

    if (!this.yearsData[year]['1']) {
      this.yearsData[year]['1'] = this.createDefaultSemesterData();
    }
    if (!this.yearsData[year]['2']) {
      const sem1 = this.yearsData[year]['1'];
      this.yearsData[year]['2'] = this.createDefaultSemesterData(sem1.periods, sem1.days, sem1.offeredSubjects, sem1.subjectColorMap);
    }

    const d = this.yearsData[year][semester];
    if (!d.periods || !Array.isArray(d.periods) || d.periods.length === 0) {
      d.periods = JSON.parse(JSON.stringify(DEFAULT_PERIODS));
    }
    if (!d.days || !Array.isArray(d.days) || d.days.length === 0) {
      d.days = JSON.parse(JSON.stringify(DEFAULT_DAYS));
    }
    if (!d.offeredSubjects || d.offeredSubjects.length === 0) {
      d.offeredSubjects = [...DEFAULT_OFFERED_SUBJECTS];
    }
    if (!d.subjectColorMap) d.subjectColorMap = {};
    if (d.showWeekends === undefined) d.showWeekends = true;

    // masterStudents & courseEnrollments & courseSchedules 마이그레이션
    const needsMigration = !d.masterStudents || !Array.isArray(d.masterStudents) || 
      (d.masterStudents.length === 0 && d.students && Array.isArray(d.students) && d.students.length > 0);

    if (needsMigration) {
      d.masterStudents = [];
      d.courseEnrollments = d.courseEnrollments || {};
      d.courseSchedules = d.courseSchedules || {};

      // 기존 d.students 로부터 자동 마이그레이션
      if (d.students && Array.isArray(d.students) && d.students.length > 0) {
        const studentMap = new Map();
        d.students.forEach(s => {
          const code = s.studentCode || ExcelService.getStudentCode(s);
          const key = `${code}_${s.name}`;
          if (!studentMap.has(key)) {
            const mStudent = {
              id: s.studentMasterId || s.id || `sm_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
              grade: s.grade || 1,
              classNum: s.classNum || 1,
              studentNum: s.studentNum || null,
              studentCode: code,
              name: s.name
            };
            studentMap.set(key, mStudent);
            d.masterStudents.push(mStudent);
          }

          const master = studentMap.get(key);
          if (s.subject) {
            if (!d.courseEnrollments[s.subject]) d.courseEnrollments[s.subject] = [];
            if (!d.courseEnrollments[s.subject].includes(master.id)) {
              d.courseEnrollments[s.subject].push(master.id);
            }

            if (s.day && (s.periodId || s.period)) {
              if (!d.courseSchedules[s.subject]) d.courseSchedules[s.subject] = [];
              const exists = d.courseSchedules[s.subject].some(sch => sch.day === s.day && (sch.periodId === s.periodId || sch.period === s.period));
              if (!exists) {
                d.courseSchedules[s.subject].push({
                  id: `sch_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
                  day: s.day,
                  periodId: s.periodId,
                  period: s.period,
                  status: s.status || 'confirmed'
                });
              }
            }
          }
        });
      }
    }
  }

  loadActiveYearSemesterData() {
    this.ensureYearSemesterData(this.activeYear, this.activeSemester);
    const d = this.yearsData[this.activeYear][this.activeSemester];
    this.periods = d.periods;
    this.days = d.days;
    this.masterStudents = d.masterStudents || [];
    this.courseEnrollments = d.courseEnrollments || {};
    this.courseSchedules = d.courseSchedules || {};
    this.offeredSubjects = d.offeredSubjects;
    this.subjectColorMap = d.subjectColorMap;
    this.showWeekends = d.showWeekends !== undefined ? d.showWeekends : true;

    // masterStudents 재확인 및 보정
    if (this.masterStudents.length === 0 && d.students && Array.isArray(d.students) && d.students.length > 0) {
      const studentMap = new Map();
      d.students.forEach(s => {
        const code = s.studentCode || ExcelService.getStudentCode(s);
        const key = `${code}_${s.name}`;
        if (!studentMap.has(key)) {
          const mStudent = {
            id: s.studentMasterId || s.id || `sm_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
            grade: s.grade || 1,
            classNum: s.classNum || 1,
            studentNum: s.studentNum || null,
            studentCode: code,
            name: s.name
          };
          studentMap.set(key, mStudent);
          this.masterStudents.push(mStudent);
        }
      });
      d.masterStudents = this.masterStudents;
    }

    this.selectedStep1StudentIds.clear();
    this.rebuildStudentAssignments();
  }

  saveActiveYearSemesterData() {
    this.sortPeriods();
    this.rebuildStudentAssignments();
    this.ensureYearSemesterData(this.activeYear, this.activeSemester);
    this.yearsData[this.activeYear][this.activeSemester] = {
      periods: this.periods,
      days: this.days,
      masterStudents: this.masterStudents,
      courseEnrollments: this.courseEnrollments,
      courseSchedules: this.courseSchedules,
      students: this.students,
      offeredSubjects: this.offeredSubjects,
      subjectColorMap: this.subjectColorMap,
      showWeekends: this.showWeekends
    };
  }

  /**
   * masterStudents, courseEnrollments, courseSchedules 정보를 기반으로
   * 시간표 및 ExcelService와 호환되는 flat `this.students` 목록을 완벽하게 재구성합니다.
   */
  rebuildStudentAssignments() {
    const list = [];
    const masterMap = new Map(this.masterStudents.map(s => [s.id, s]));

    // 1. 과목별 편성 및 시간표 슬롯 매핑
    this.offeredSubjects.forEach(subject => {
      const studentIds = this.courseEnrollments[subject] || [];
      const schedules = this.courseSchedules[subject] || [];

      if (schedules.length === 0) {
        // 시간표에 아직 배치되지 않은 수강생들 (미배치 상태로 등록)
        studentIds.forEach(stId => {
          const m = masterMap.get(stId);
          if (m) {
            list.push({
              id: `assign_${subject}_unassigned_${m.id}`,
              studentMasterId: m.id,
              grade: m.grade,
              classNum: m.classNum,
              studentNum: m.studentNum,
              studentCode: m.studentCode || ExcelService.getStudentCode(m),
              name: m.name,
              subject: subject,
              day: null,
              periodId: null,
              period: null,
              status: 'confirmed'
            });
          }
        });
      } else {
        // 시간표에 1개 이상의 요일/교시로 배치된 경우
        schedules.forEach(slot => {
          const periodObj = this.periods.find(p => p.id === slot.periodId || (!slot.periodId && p.period === slot.period));
          studentIds.forEach(stId => {
            const m = masterMap.get(stId);
            if (m) {
              list.push({
                id: `assign_${subject}_${slot.id || (slot.day + '_' + slot.periodId)}_${m.id}`,
                scheduleId: slot.id,
                studentMasterId: m.id,
                grade: m.grade,
                classNum: m.classNum,
                studentNum: m.studentNum,
                studentCode: m.studentCode || ExcelService.getStudentCode(m),
                name: m.name,
                subject: subject,
                day: slot.day,
                periodId: slot.periodId,
                period: periodObj ? periodObj.period : slot.period,
                status: slot.status || 'confirmed'
              });
            }
          });
        });
      }
    });

    // 2. 어떤 과목에도 아직 편제되지 않은 순수 미편제 학생 등록
    const enrolledStudentIds = new Set();
    Object.values(this.courseEnrollments).forEach(arr => {
      arr.forEach(id => enrolledStudentIds.add(id));
    });

    this.masterStudents.forEach(m => {
      if (!enrolledStudentIds.has(m.id)) {
        list.push({
          id: `master_pure_${m.id}`,
          studentMasterId: m.id,
          grade: m.grade,
          classNum: m.classNum,
          studentNum: m.studentNum,
          studentCode: m.studentCode || ExcelService.getStudentCode(m),
          name: m.name,
          subject: '',
          day: null,
          periodId: null,
          period: null,
          status: 'confirmed'
        });
      }
    });

    this.students = list;
  }

  loadState() {
    try {
      let saved = localStorage.getItem(this.storageKey);
      if (!saved) {
        // v3 레거시 키 확인
        saved = localStorage.getItem(this.legacyStorageKey);
      }

      if (saved) {
        const data = JSON.parse(saved);
        if (data.yearsData || data.years) {
          this.yearsData = data.yearsData || data.years;
          this.academicYears = data.academicYears || Object.keys(this.yearsData);
          if (!this.academicYears.includes('2026')) this.academicYears.unshift('2026');
          if (!this.academicYears.includes('2027')) this.academicYears.push('2027');
          this.activeYear = data.activeYear || '2026';
          this.activeSemester = data.activeSemester || '1';
        } else {
          this.academicYears = ['2026', '2027'];
          this.activeYear = '2026';
          this.activeSemester = '1';
          this.yearsData['2026'] = {
            '1': {
              periods: (data.periods && data.periods.length > 0) ? data.periods : JSON.parse(JSON.stringify(DEFAULT_PERIODS)),
              days: (data.days && data.days.length > 0) ? data.days : JSON.parse(JSON.stringify(DEFAULT_DAYS)),
              students: data.students || [],
              offeredSubjects: (data.offeredSubjects && data.offeredSubjects.length > 0) ? data.offeredSubjects : [...DEFAULT_OFFERED_SUBJECTS],
              subjectColorMap: data.subjectColorMap || {},
              showWeekends: data.showWeekends !== undefined ? data.showWeekends : true
            },
            '2': this.createDefaultSemesterData()
          };
          this.yearsData['2027'] = {
            '1': this.createDefaultSemesterData(),
            '2': this.createDefaultSemesterData()
          };
        }

        this.ensureYearSemesterData('2026', '1');
        this.ensureYearSemesterData('2026', '2');
        this.ensureYearSemesterData('2027', '1');
        this.ensureYearSemesterData('2027', '2');
        this.loadActiveYearSemesterData();
        return;
      }
    } catch (e) {
      console.error('Failed to load state:', e);
    }

    this.resetToDefaults(false);
  }

  saveState() {
    try {
      this.saveActiveYearSemesterData();
      const data = {
        activeYear: this.activeYear,
        activeSemester: this.activeSemester,
        academicYears: this.academicYears,
        yearsData: this.yearsData,
        periods: this.periods,
        days: this.days,
        students: this.students,
        masterStudents: this.masterStudents,
        courseEnrollments: this.courseEnrollments,
        courseSchedules: this.courseSchedules,
        offeredSubjects: this.offeredSubjects,
        subjectColorMap: this.subjectColorMap,
        showWeekends: this.showWeekends,
        lastUpdated: new Date().toISOString()
      };
      localStorage.setItem(this.storageKey, JSON.stringify(data));
      this.updateStats();
      this.renderAcademicYearTabs();

      if (window.cloudSync) {
        window.cloudSync.saveToBackend(data);
      }
    } catch (e) {
      console.error('Failed to save state:', e);
    }
  }

  applyExternalData(data, msg = '') {
    if (!data) return;
    if (data.yearsData || data.years) {
      this.yearsData = data.yearsData || data.years;
      this.academicYears = data.academicYears || Object.keys(this.yearsData);
      if (!this.academicYears.includes('2026')) this.academicYears.unshift('2026');
      if (!this.academicYears.includes('2027')) this.academicYears.push('2027');
      if (data.activeYear && this.academicYears.includes(data.activeYear)) this.activeYear = data.activeYear;
      if (data.activeSemester) this.activeSemester = data.activeSemester;
    }

    this.loadActiveYearSemesterData();
    this.sortPeriods();
    this.buildSubjectColorMap();
    this.saveState();
    this.render();
    if (msg) this.showToast(msg);
  }

  resetToDefaults(showConfirm = true) {
    if (showConfirm && !confirm(`현재 [${this.activeYear}학년도 ${this.activeSemester}학기] 데이터를 초기화하시겠습니까?`)) {
      return;
    }
    this.ensureYearSemesterData(this.activeYear, this.activeSemester);
    this.yearsData[this.activeYear][this.activeSemester] = this.createDefaultSemesterData();
    this.loadActiveYearSemesterData();
    this.buildSubjectColorMap();
    this.saveState();
    this.render();
    this.showToast('데이터가 초기화되었습니다.');
  }

  // ==========================================
  // 워크플로우 탭 전환 (Step 1 vs Step 2)
  // ==========================================
  switchWorkflowTab(tabName) {
    this.currentWorkflowTab = tabName;

    const btn1 = document.getElementById('tab-btn-step1');
    const btn2 = document.getElementById('tab-btn-step2');
    const view1 = document.getElementById('view-step1');
    const view2 = document.getElementById('view-step2');
    const hintText = document.getElementById('workflow-hint-text');

    if (tabName === 'step1') {
      if (view1) view1.classList.remove('hidden');
      if (view2) view2.classList.add('hidden');
      if (btn1) {
        btn1.className = 'px-4 py-2 rounded-lg text-xs font-bold flex items-center gap-2 transition bg-blue-600 text-white shadow-xs';
      }
      if (btn2) {
        btn2.className = 'px-4 py-2 rounded-lg text-xs font-bold flex items-center gap-2 transition bg-slate-100 hover:bg-slate-200 text-slate-700';
      }
      if (hintText) {
        hintText.innerHTML = '학생을 <strong>과목 바구니로 드래그</strong>하여 편제하세요. 과목 내에서 <strong>시간을 추가</strong>하면 시간표에 자동 배치됩니다.';
      }
    } else {
      if (view1) view1.classList.add('hidden');
      if (view2) view2.classList.remove('hidden');
      if (btn1) {
        btn1.className = 'px-4 py-2 rounded-lg text-xs font-bold flex items-center gap-2 transition bg-slate-100 hover:bg-slate-200 text-slate-700';
      }
      if (btn2) {
        btn2.className = 'px-4 py-2 rounded-lg text-xs font-bold flex items-center gap-2 transition bg-blue-600 text-white shadow-xs';
      }
      if (hintText) {
        hintText.innerHTML = '과목 카드를 시간표 셀로 <strong>드래그</strong>하여 배치하세요. (카드를 그냥 드래그: <strong>이동</strong> | Ctrl+드래그: <strong>복사</strong>)';
      }
    }

    this.render();
  }

  // ==========================================
  // 학년도 및 학기 전환
  // ==========================================
  switchAcademicYear(year) {
    if (this.activeYear === year) return;
    this.saveActiveYearSemesterData();
    this.activeYear = year;
    this.loadActiveYearSemesterData();
    this.sortPeriods();
    this.buildSubjectColorMap();
    this.saveState();
    this.render();
    this.showToast(`[${this.activeYear}학년도 ${this.activeSemester}학기] 시간표로 전환되었습니다.`);
  }

  switchSemester(semester) {
    if (this.activeSemester === semester) return;
    this.saveActiveYearSemesterData();
    this.activeSemester = semester;
    this.loadActiveYearSemesterData();
    this.sortPeriods();
    this.buildSubjectColorMap();
    this.saveState();
    this.render();
    this.showToast(`[${this.activeYear}학년도 ${this.activeSemester}학기] 시간표로 전환되었습니다.`);
  }

  handleAddAcademicYear() {
    const defaultNext = String(Math.max(...this.academicYears.map(Number)) + 1);
    const inputYear = prompt('추가할 학년도를 4자리 숫자로 입력하세요 (예: 2028):', defaultNext);
    if (!inputYear) return;
    const cleanYear = inputYear.trim().replace(/[^0-9]/g, '');
    if (!cleanYear || cleanYear.length !== 4) {
      alert('올바른 4자리 연도를 입력해주세요 (예: 2028)');
      return;
    }
    if (this.academicYears.includes(cleanYear)) {
      alert(`'${cleanYear}학년도'는 이미 등록되어 있습니다.`);
      this.switchAcademicYear(cleanYear);
      return;
    }

    this.saveActiveYearSemesterData();
    this.academicYears.push(cleanYear);
    this.academicYears.sort((a, b) => Number(a) - Number(b));
    this.yearsData[cleanYear] = {
      '1': this.createDefaultSemesterData(this.periods, this.days, this.offeredSubjects, this.subjectColorMap),
      '2': this.createDefaultSemesterData(this.periods, this.days, this.offeredSubjects, this.subjectColorMap)
    };
    this.activeYear = cleanYear;
    this.activeSemester = '1';
    this.loadActiveYearSemesterData();
    this.saveState();
    this.render();
    this.showToast(`'${cleanYear}학년도' 탭이 추가되었습니다.`);
  }

  handleDeleteAcademicYear(yearToDelete) {
    if (this.academicYears.length <= 1) {
      alert('최소 1개 이상의 학년도가 유지되어야 합니다.');
      return;
    }
    if (!confirm(`'${yearToDelete}학년도' 탭과 해당 학년도의 모든 데이터를 정말 삭제하시겠습니까?`)) {
      return;
    }

    delete this.yearsData[yearToDelete];
    this.academicYears = this.academicYears.filter(y => y !== yearToDelete);
    if (this.activeYear === yearToDelete) {
      this.activeYear = this.academicYears[0];
    }
    this.loadActiveYearSemesterData();
    this.saveState();
    this.render();
    this.showToast(`'${yearToDelete}학년도'가 삭제되었습니다.`);
  }

  // ==========================================
  // 일과표 교시 정렬
  // ==========================================
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
    const subSet = new Set(this.offeredSubjects);
    Object.keys(this.courseEnrollments || {}).forEach(s => subSet.add(s));
    this.offeredSubjects = Array.from(subSet).filter(Boolean);
  }

  buildSubjectColorMap() {
    this.offeredSubjects.forEach((sub, idx) => {
      if (!this.subjectColorMap[sub]) {
        this.subjectColorMap[sub] = PASTEL_PALETTE[idx % PASTEL_PALETTE.length];
      }
    });
  }

  getSubjectColor(subject) {
    if (!subject) return { bg: '#F1F5F9', text: '#475569', border: '#CBD5E1', excelBg: 'F1F5F9', excelText: '475569' };
    if (!this.subjectColorMap[subject]) {
      const idx = Object.keys(this.subjectColorMap).length;
      this.subjectColorMap[subject] = PASTEL_PALETTE[idx % PASTEL_PALETTE.length];
    }
    return this.subjectColorMap[subject];
  }

  // ==========================================
  // 이벤트 바인딩
  // ==========================================
  bindEvents() {
    // 1. 키보드 Ctrl/Cmd 키 추적 (복사 모드 토글)
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Control' || e.key === 'Meta') {
        this.isCtrlPressed = true;
        const banner = document.getElementById('ctrl-copy-indicator');
        if (banner) banner.classList.remove('hidden');
        document.body.classList.add('is-copy-mode');
      }
    });

    window.addEventListener('keyup', (e) => {
      if (e.key === 'Control' || e.key === 'Meta') {
        this.isCtrlPressed = false;
        const banner = document.getElementById('ctrl-copy-indicator');
        if (banner) banner.classList.add('hidden');
        document.body.classList.remove('is-copy-mode');
      }
    });

    // 2. 탭 1 퀵 학생 직접 등록
    document.getElementById('form-quick-add-student')?.addEventListener('submit', (e) => {
      e.preventDefault();
      this.handleQuickAddStudent();
    });

    // 3. 탭 1 새 과목 바구니 생성
    document.getElementById('form-create-subject-bucket')?.addEventListener('submit', (e) => {
      e.preventDefault();
      this.handleCreateSubjectBucket();
    });

    // 4. 탭 1 학생 다중 선택 & 일괄 바구니 담기
    document.getElementById('step1-select-all')?.addEventListener('change', (e) => {
      this.handleToggleSelectAllStudents(e.target.checked);
    });

    document.getElementById('step1-btn-batch-enroll')?.addEventListener('click', () => {
      this.handleBatchEnrollSelectedStudents();
    });

    // 5. 탭 1 검색 및 필터
    document.getElementById('step1-search-input')?.addEventListener('input', (e) => {
      this.step1Filter.search = e.target.value.trim().toLowerCase();
      this.renderStep1StudentList();
    });

    document.getElementById('step1-filter-grade')?.addEventListener('change', (e) => {
      this.step1Filter.grade = e.target.value;
      this.renderStep1StudentList();
    });

    document.getElementById('step1-filter-enroll')?.addEventListener('change', (e) => {
      this.step1Filter.enroll = e.target.value;
      this.renderStep1StudentList();
    });

    // 6. 엑셀 학생 업로드
    const handleStudentUpload = async (file) => {
      if (!file) return;
      try {
        const newStudents = await ExcelService.parseStudentExcel(file);
        const action = confirm(
          `총 ${newStudents.length}명의 학생을 불러왔습니다!\n\n` +
          `[확인]: 기존 학생 목록을 유지하고 새로 불러온 학생 추가\n` +
          `[취소]: 작업 취소`
        );
        if (action) {
          const seen = new Set(this.masterStudents.map(s => `${s.studentCode}_${s.name}`));
          let added = 0;
          newStudents.forEach(st => {
            const code = st.studentCode || ExcelService.getStudentCode(st);
            const key = `${code}_${st.name}`;
            if (!seen.has(key)) {
              seen.add(key);
              this.masterStudents.push({
                id: `sm_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
                grade: st.grade,
                classNum: st.classNum,
                studentNum: st.studentNum,
                studentCode: code,
                name: st.name
              });
              added++;
            }
          });
          this.rebuildStudentAssignments();
          this.saveState();
          this.render();
          this.showToast(`${added}명의 학생이 새로 등록되었습니다.`);
        }
      } catch (err) {
        alert('학생 명단 파일 처리 실패: ' + err.message);
      }
    };

    document.getElementById('file-upload-student')?.addEventListener('change', (e) => {
      handleStudentUpload(e.target.files[0]);
      e.target.value = '';
    });
    document.getElementById('step1-file-upload-student')?.addEventListener('change', (e) => {
      handleStudentUpload(e.target.files[0]);
      e.target.value = '';
    });

    // 7. 교시 추가 모달
    document.getElementById('btn-add-period')?.addEventListener('click', () => {
      this.openAddPeriodModal();
    });
    document.getElementById('btn-close-period-modal')?.addEventListener('click', () => {
      this.closeAddPeriodModal();
    });
    document.getElementById('form-add-period')?.addEventListener('submit', (e) => {
      e.preventDefault();
      this.handleSaveCustomPeriod();
    });

    // 8. 주말 표시 토글
    document.getElementById('btn-toggle-weekends')?.addEventListener('click', () => {
      this.showWeekends = !this.showWeekends;
      this.saveState();
      this.renderTimetable();
      this.updateWeekendToggleBtn();
    });

    // 9. 엑셀 다운로드 (학급 시간표 & 원천 데이터)
    document.getElementById('btn-export-timetable')?.addEventListener('click', () => {
      const activeDays = this.showWeekends ? this.days : this.days.filter(d => !d.isWeekend);
      ExcelService.exportTimetableExcel(this.periods, activeDays, this.students, this.subjectColorMap, this.activeYear, this.activeSemester);
      this.showToast(`[${this.activeYear}학년도 ${this.activeSemester}학기] 학급 시간표 엑셀이 다운로드됩니다.`);
    });

    document.getElementById('btn-export-data')?.addEventListener('click', () => {
      const activeDays = this.showWeekends ? this.days : this.days.filter(d => !d.isWeekend);
      ExcelService.exportDataExcel(this.students, this.periods, activeDays, this.activeYear, this.activeSemester);
      this.showToast(`[${this.activeYear}학년도 ${this.activeSemester}학기] 원천 데이터 엑셀이 다운로드됩니다.`);
    });

    // 10. 일과표 업로드
    document.getElementById('file-upload-period')?.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      try {
        const newPeriods = await ExcelService.parsePeriodExcel(file);
        if (confirm(`총 ${newPeriods.length}개의 교시 일과표 데이터를 적용하시겠습니까?`)) {
          this.periods = newPeriods;
          this.sortPeriods();
          this.rebuildStudentAssignments();
          this.saveState();
          this.render();
          this.showToast('일과표가 성공적으로 적용되었습니다.');
        }
      } catch (err) {
        alert('일과표 파일 처리 실패: ' + err.message);
      }
      e.target.value = '';
    });

    // 11. 양식 다운로드
    document.getElementById('btn-dl-period-template')?.addEventListener('click', () => {
      ExcelService.downloadPeriodTemplate();
    });
    document.getElementById('btn-dl-student-template')?.addEventListener('click', () => {
      ExcelService.downloadStudentTemplate();
    });

    // 12. 인쇄
    document.getElementById('btn-print')?.addEventListener('click', () => {
      window.print();
    });

    // 13. 상세 학생 등록 모달
    document.getElementById('form-add-student')?.addEventListener('submit', (e) => {
      e.preventDefault();
      this.handleSaveDetailedStudent();
    });
    document.getElementById('btn-close-modal')?.addEventListener('click', () => {
      this.closeStudentModal();
    });

    // 14. 새 학년도 추가 버튼
    document.getElementById('btn-add-academic-year')?.addEventListener('click', () => {
      this.handleAddAcademicYear();
    });

    // 15. 초기화
    document.getElementById('btn-reset-data')?.addEventListener('click', () => {
      this.resetToDefaults(true);
    });
  }

  // ==========================================
  // 렌더링 총괄
  // ==========================================
  render() {
    this.updateStats();
    this.renderAcademicYearTabs();
    this.updateWeekendToggleBtn();

    if (this.currentWorkflowTab === 'step1') {
      this.renderStep1StudentList();
      this.renderSubjectBuckets();
      this.populateBatchSubjectSelect();
    } else {
      this.renderStep2Palette();
      this.renderTimetable();
    }

    if (window.lucide) window.lucide.createIcons();
  }

  updateStats() {
    const statTotal = document.getElementById('stat-total-students');
    const statSubjects = document.getElementById('stat-subject-count');
    const statAssigned = document.getElementById('stat-assigned-students');
    const statUnassigned = document.getElementById('stat-unassigned-students');
    const statPeriod = document.getElementById('stat-period-count');

    const totalStudentsCount = this.masterStudents.length;
    const subjectCount = this.offeredSubjects.length;

    // 배정 완료된 고유 학생 수 (시간표에 최소 1교시 이상 편성된 학생)
    const scheduledStudentIds = new Set();
    Object.entries(this.courseSchedules).forEach(([sub, scheds]) => {
      if (scheds.length > 0) {
        const enrolled = this.courseEnrollments[sub] || [];
        enrolled.forEach(id => scheduledStudentIds.add(id));
      }
    });

    const assignedCount = scheduledStudentIds.size;
    const unassignedCount = Math.max(0, totalStudentsCount - assignedCount);
    const regularPeriodsCount = this.periods.filter(p => !p.isBreak).length;

    if (statTotal) statTotal.textContent = totalStudentsCount;
    if (statSubjects) statSubjects.textContent = subjectCount;
    if (statAssigned) statAssigned.textContent = assignedCount;
    if (statUnassigned) statUnassigned.textContent = unassignedCount;
    if (statPeriod) statPeriod.textContent = regularPeriodsCount;
  }

  renderAcademicYearTabs() {
    const container = document.getElementById('academic-year-tabs');
    const badgeLabel = document.getElementById('current-year-label');
    const titleYear = document.getElementById('timetable-academic-year-text');
    const printYear = document.getElementById('timetable-print-year-text');

    const fullYearSemText = `${this.activeYear}학년도 ${this.activeSemester}학기`;
    if (badgeLabel) badgeLabel.textContent = fullYearSemText;
    if (titleYear) titleYear.textContent = fullYearSemText;
    if (printYear) printYear.textContent = fullYearSemText;

    if (container) {
      container.innerHTML = this.academicYears.map(year => {
        const isActive = year === this.activeYear;
        const yearObj = this.yearsData[year] || {};
        const count1 = (yearObj['1'] && yearObj['1'].masterStudents) ? yearObj['1'].masterStudents.length : 0;
        const count2 = (yearObj['2'] && yearObj['2'].masterStudents) ? yearObj['2'].masterStudents.length : 0;
        const totalCount = count1 + count2;

        return `
          <div class="inline-flex items-center rounded-lg transition-all ${
            isActive 
              ? 'bg-blue-600 text-white shadow-xs font-bold ring-2 ring-blue-400/40' 
              : 'bg-white hover:bg-slate-50 text-slate-700 font-semibold border border-slate-300'
          } text-xs">
            <button type="button" 
                    onclick="window.app.switchAcademicYear('${year}')" 
                    class="px-3 py-1.5 flex items-center gap-1.5 focus:outline-none"
                    title="${year}학년도로 전환">
              <i data-lucide="${isActive ? 'calendar-check' : 'calendar'}" class="w-3.5 h-3.5 ${isActive ? 'text-white' : 'text-slate-400'}"></i>
              <span>${year}학년도</span>
              ${totalCount > 0 ? `
                <span class="text-[10px] px-1.5 py-0.2 rounded-full font-bold ${
                  isActive ? 'bg-blue-500 text-white' : 'bg-slate-100 text-slate-600'
                }">
                  ${totalCount}명
                </span>
              ` : ''}
            </button>
            ${this.academicYears.length > 1 ? `
              <button type="button" 
                      onclick="event.stopPropagation(); window.app.handleDeleteAcademicYear('${year}')" 
                      class="pr-2 pl-0.5 opacity-60 hover:opacity-100 ${isActive ? 'hover:text-red-200 text-white' : 'hover:text-red-600 text-slate-400'} transition" 
                      title="'${year}학년도' 삭제">
                &times;
              </button>
            ` : ''}
          </div>
        `;
      }).join('');
    }

    const currentYearObj = this.yearsData[this.activeYear] || {};
    const sem1Count = (currentYearObj['1'] && currentYearObj['1'].masterStudents) ? currentYearObj['1'].masterStudents.length : 0;
    const sem2Count = (currentYearObj['2'] && currentYearObj['2'].masterStudents) ? currentYearObj['2'].masterStudents.length : 0;

    const btnSem1 = document.getElementById('tab-semester-1');
    const btnSem2 = document.getElementById('tab-semester-2');
    const badgeSem1 = document.getElementById('badge-sem-1-count');
    const badgeSem2 = document.getElementById('badge-sem-2-count');

    if (btnSem1 && btnSem2) {
      if (this.activeSemester === '1') {
        btnSem1.className = 'px-3.5 py-1 text-xs font-bold rounded-md bg-blue-600 text-white shadow-xs flex items-center gap-1.5 transition';
        btnSem2.className = 'px-3.5 py-1 text-xs font-semibold rounded-md text-slate-700 hover:bg-slate-300/60 flex items-center gap-1.5 transition';
      } else {
        btnSem1.className = 'px-3.5 py-1 text-xs font-semibold rounded-md text-slate-700 hover:bg-slate-300/60 flex items-center gap-1.5 transition';
        btnSem2.className = 'px-3.5 py-1 text-xs font-bold rounded-md bg-blue-600 text-white shadow-xs flex items-center gap-1.5 transition';
      }
    }

    if (badgeSem1) badgeSem1.textContent = `${sem1Count}명`;
    if (badgeSem2) badgeSem2.textContent = `${sem2Count}명`;
  }

  updateWeekendToggleBtn() {
    const btn = document.getElementById('btn-toggle-weekends');
    if (!btn) return;
    btn.innerHTML = `
      <i data-lucide="calendar" class="w-4 h-4"></i>
      <span>${this.showWeekends ? '주말 숨기기(월~금)' : '주말 포함(월~일)'}</span>
    `;
    if (window.lucide) window.lucide.createIcons();
  }

  // ==========================================
  // [1단계: 탭 1] 학생 등록 및 목록 (학생 풀)
  // ==========================================
  handleQuickAddStudent() {
    const grade = parseInt(document.getElementById('quick-grade').value, 10) || 1;
    const classNum = parseInt(document.getElementById('quick-class').value, 10) || 1;
    const studentNum = parseInt(document.getElementById('quick-number').value, 10) || 1;
    const nameEl = document.getElementById('quick-name');
    const name = nameEl.value.trim();

    if (!name) {
      alert('학생 성명을 입력해주세요.');
      nameEl.focus();
      return;
    }

    const studentCode = ExcelService.getStudentCode({ grade, classNum, studentNum });
    const exists = this.masterStudents.some(s => (s.studentCode || ExcelService.getStudentCode(s)) === studentCode && s.name === name);
    if (exists) {
      alert(`학번 ${studentCode} ${name} 학생은 이미 명단에 등록되어 있습니다.`);
      return;
    }

    const newStudent = {
      id: `sm_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
      grade,
      classNum,
      studentNum,
      studentCode,
      name
    };

    this.masterStudents.push(newStudent);
    this.rebuildStudentAssignments();
    this.saveState();
    this.render();

    nameEl.value = '';
    // 다음 번호 자동 증가
    document.getElementById('quick-number').value = studentNum + 1;
    nameEl.focus();
    this.showToast(`[${studentCode} ${name}] 학생이 등록되었습니다.`);
  }

  handleSaveDetailedStudent() {
    const grade = parseInt(document.getElementById('modal-grade').value, 10) || 1;
    const classNum = parseInt(document.getElementById('modal-class').value, 10) || 1;
    const studentNum = parseInt(document.getElementById('modal-number').value, 10) || null;
    const name = document.getElementById('modal-name').value.trim();
    const targetSubject = document.getElementById('modal-subject').value.trim();

    if (!name) {
      alert('학생 성명을 입력해주세요.');
      return;
    }

    const studentCode = ExcelService.getStudentCode({ grade, classNum, studentNum });
    const newStudent = {
      id: `sm_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
      grade,
      classNum,
      studentNum,
      studentCode,
      name
    };

    this.masterStudents.push(newStudent);

    if (targetSubject) {
      if (!this.offeredSubjects.includes(targetSubject)) {
        this.offeredSubjects.push(targetSubject);
        this.buildSubjectColorMap();
      }
      if (!this.courseEnrollments[targetSubject]) this.courseEnrollments[targetSubject] = [];
      this.courseEnrollments[targetSubject].push(newStudent.id);
    }

    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.closeStudentModal();
    this.showToast(`'${studentCode} ${name}' 학생이 등록되었습니다.`);
  }

  getFilteredMasterStudents() {
    return this.masterStudents.filter(student => {
      // 1. 학년 필터
      if (this.step1Filter.grade !== 'all') {
        if (String(student.grade) !== String(this.step1Filter.grade)) return false;
      }

      // 2. 편제 상태 필터
      const enrolledCourses = this.getStudentEnrolledCourses(student.id);
      const isEnrolled = enrolledCourses.length > 0;
      if (this.step1Filter.enroll === 'unassigned' && isEnrolled) return false;
      if (this.step1Filter.enroll === 'assigned' && !isEnrolled) return false;

      // 3. 검색어 필터
      if (this.step1Filter.search) {
        const q = this.step1Filter.search;
        const code = (student.studentCode || ExcelService.getStudentCode(student)).toLowerCase();
        const name = (student.name || '').toLowerCase();
        const hasSub = enrolledCourses.some(c => c.toLowerCase().includes(q));
        if (!code.includes(q) && !name.includes(q) && !hasSub) return false;
      }

      return true;
    });
  }

  getStudentEnrolledCourses(studentId) {
    const list = [];
    Object.entries(this.courseEnrollments).forEach(([subject, ids]) => {
      if (ids.includes(studentId)) list.push(subject);
    });
    return list;
  }

  renderStep1StudentList() {
    const container = document.getElementById('step1-student-list-container');
    const badgeCount = document.getElementById('step1-student-count-badge');
    if (!container) return;

    const filtered = this.getFilteredMasterStudents();
    if (badgeCount) badgeCount.textContent = `${this.masterStudents.length}명`;

    if (filtered.length === 0) {
      container.innerHTML = `
        <div class="py-12 text-center text-slate-400 text-xs">
          <i data-lucide="users" class="w-8 h-8 mx-auto mb-2 opacity-40"></i>
          <p class="font-medium text-slate-600">등록된 학생이 없습니다.</p>
          <p class="text-[11px] text-slate-400 mt-1">상단의 <strong>[학생 직접 등록]</strong> 또는 <strong>[엑셀 업로드]</strong>로 추가하세요.</p>
        </div>
      `;
      if (window.lucide) window.lucide.createIcons();
      return;
    }

    let html = '';
    filtered.forEach(st => {
      const code = st.studentCode || ExcelService.getStudentCode(st);
      const isSelected = this.selectedStep1StudentIds.has(st.id);
      const enrolledCourses = this.getStudentEnrolledCourses(st.id);

      html += `
        <div class="draggable-student-item p-2.5 rounded-lg border transition-all ${
          isSelected 
            ? 'bg-blue-50/80 border-blue-300 shadow-xs ring-1 ring-blue-400/40' 
            : 'bg-white border-slate-200 hover:border-blue-300'
        } flex items-center justify-between gap-2"
             draggable="true"
             ondragstart="window.app.handleStudentDragStart(event, '${st.id}')"
             ondragend="window.app.handleStudentDragEnd(event)">
          
          <div class="flex items-center gap-2 min-w-0">
            <!-- 체크박스 -->
            <input type="checkbox" value="${st.id}" ${isSelected ? 'checked' : ''}
                   onchange="window.app.toggleSelectStep1Student('${st.id}', this.checked)"
                   class="rounded text-blue-600 focus:ring-blue-500">
            
            <!-- 드래그 손잡이 아이콘 -->
            <i data-lucide="grip-vertical" class="w-3.5 h-3.5 text-slate-400 shrink-0"></i>

            <!-- 학번 & 성명 -->
            <div class="flex items-baseline gap-1.5 min-w-0">
              <span class="font-mono font-bold text-xs text-slate-900">${code}</span>
              <span class="font-bold text-xs text-slate-800 truncate">${st.name}</span>
            </div>
          </div>

          <!-- 편제된 과목 뱃지들 -->
          <div class="flex items-center gap-1 shrink-0 flex-wrap justify-end max-w-[190px]">
            ${enrolledCourses.length > 0 ? enrolledCourses.map(sub => {
              const color = this.getSubjectColor(sub);
              return `
                <span class="text-[10px] font-bold px-1.5 py-0.5 rounded border flex items-center gap-1"
                      style="background-color: ${color.bg}; color: ${color.text}; border-color: ${color.border};">
                  <span>${sub}</span>
                  <button type="button" 
                          onclick="event.stopPropagation(); window.app.handleRemoveStudentFromCourse('${sub}', '${st.id}')"
                          class="hover:text-red-600 opacity-60 hover:opacity-100" title="'${sub}' 과목에서 제외">&times;</button>
                </span>
              `;
            }).join('') : `
              <span class="text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-400 font-medium">미편제</span>
            `}

            <!-- 학생 삭제 버튼 -->
            <button type="button" 
                    onclick="event.stopPropagation(); window.app.deleteMasterStudent('${st.id}')" 
                    class="p-1 text-slate-300 hover:text-red-500 rounded transition ml-1" 
                    title="학생 명단에서 완전 삭제">
              <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
            </button>
          </div>

        </div>
      `;
    });

    container.innerHTML = html;
    if (window.lucide) window.lucide.createIcons();
    this.updateStep1SelectionUI();
  }

  handleStudentDragStart(e, studentId) {
    const isSelected = this.selectedStep1StudentIds.has(studentId);
    const studentIds = isSelected 
      ? Array.from(this.selectedStep1StudentIds) 
      : [studentId];

    this.dragData = {
      type: 'student',
      studentId: studentId,
      studentIds: studentIds
    };

    e.dataTransfer.setData('text/plain', JSON.stringify(this.dragData));
    e.dataTransfer.effectAllowed = 'copyMove';
    e.currentTarget.classList.add('is-dragging');
  }

  handleStudentDragEnd(e) {
    this.dragData = null;
    e.currentTarget.classList.remove('is-dragging');
    document.querySelectorAll('.bucket-drop-zone.drag-over').forEach(el => el.classList.remove('drag-over'));
  }

  toggleSelectStep1Student(studentId, isChecked) {
    if (isChecked) {
      this.selectedStep1StudentIds.add(studentId);
    } else {
      this.selectedStep1StudentIds.delete(studentId);
    }
    this.updateStep1SelectionUI();
  }

  handleToggleSelectAllStudents(isChecked) {
    const filtered = this.getFilteredMasterStudents();
    if (isChecked) {
      filtered.forEach(s => this.selectedStep1StudentIds.add(s.id));
    } else {
      filtered.forEach(s => this.selectedStep1StudentIds.delete(s.id));
    }
    this.renderStep1StudentList();
  }

  updateStep1SelectionUI() {
    const count = this.selectedStep1StudentIds.size;
    const badge = document.getElementById('step1-selected-count');
    const selectAllBox = document.getElementById('step1-select-all');
    const btnEnroll = document.getElementById('step1-btn-batch-enroll');

    if (badge) badge.textContent = `${count}명 선택`;
    if (btnEnroll) btnEnroll.disabled = count === 0;

    const filtered = this.getFilteredMasterStudents();
    if (selectAllBox && filtered.length > 0) {
      selectAllBox.checked = filtered.every(s => this.selectedStep1StudentIds.has(s.id));
    }
  }

  populateBatchSubjectSelect() {
    const sel = document.getElementById('step1-batch-subject-select');
    if (!sel) return;
    sel.innerHTML = this.offeredSubjects.map(sub => {
      return `<option value="${sub}">${sub}</option>`;
    }).join('');
  }

  handleBatchEnrollSelectedStudents() {
    if (this.selectedStep1StudentIds.size === 0) {
      alert('과목에 담을 학생을 먼저 체크해주세요.');
      return;
    }
    const select = document.getElementById('step1-batch-subject-select');
    const targetSubject = select.value;
    if (!targetSubject) {
      alert('과목을 선택해주세요.');
      return;
    }

    if (!this.courseEnrollments[targetSubject]) this.courseEnrollments[targetSubject] = [];
    let addedCount = 0;
    this.selectedStep1StudentIds.forEach(id => {
      if (!this.courseEnrollments[targetSubject].includes(id)) {
        this.courseEnrollments[targetSubject].push(id);
        addedCount++;
      }
    });

    this.selectedStep1StudentIds.clear();
    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`선택한 ${addedCount}명의 학생이 '${targetSubject}' 바구니에 담겼습니다.`);
  }

  deleteMasterStudent(studentId) {
    const st = this.masterStudents.find(s => s.id === studentId);
    if (!st) return;
    if (!confirm(`'${st.studentCode} ${st.name}' 학생을 명단에서 삭제하시겠습니까?\n(모든 과목 편제 및 시간표에서도 제외됩니다)`)) {
      return;
    }

    this.masterStudents = this.masterStudents.filter(s => s.id !== studentId);
    this.selectedStep1StudentIds.delete(studentId);

    // 모든 과목 편제에서도 제외
    Object.keys(this.courseEnrollments).forEach(sub => {
      this.courseEnrollments[sub] = this.courseEnrollments[sub].filter(id => id !== studentId);
    });

    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`${st.name} 학생이 삭제되었습니다.`);
  }

  // ==========================================
  // [1단계: 탭 1] 과목 바구니 (Folder/Bucket)
  // ==========================================
  handleCreateSubjectBucket() {
    const input = document.getElementById('input-new-bucket-name');
    const name = input.value.trim();
    if (!name) return;

    if (this.offeredSubjects.includes(name)) {
      alert(`'${name}' 과목 바구니가 이미 존재합니다.`);
      input.focus();
      return;
    }

    this.offeredSubjects.push(name);
    if (!this.courseEnrollments[name]) this.courseEnrollments[name] = [];
    if (!this.courseSchedules[name]) this.courseSchedules[name] = [];

    this.buildSubjectColorMap();
    this.saveState();
    this.render();

    input.value = '';
    this.showToast(`'${name}' 과목 바구니가 생성되었습니다. 학생을 드래그하여 편제하세요!`);
  }

  handleDeleteSubjectBucket(subjectName) {
    if (!confirm(`'${subjectName}' 과목 바구니를 삭제하시겠습니까?\n(편제된 학생은 학생 풀로 되돌아가며, 시간표에서도 제외됩니다)`)) {
      return;
    }

    this.offeredSubjects = this.offeredSubjects.filter(s => s !== subjectName);
    delete this.courseEnrollments[subjectName];
    delete this.courseSchedules[subjectName];

    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`'${subjectName}' 과목 바구니가 삭제되었습니다.`);
  }

  renderSubjectBuckets() {
    const container = document.getElementById('subject-buckets-container');
    if (!container) return;

    if (this.offeredSubjects.length === 0) {
      container.innerHTML = `
        <div class="col-span-full py-16 text-center text-slate-400 bg-slate-50/60 rounded-xl border border-dashed border-slate-300">
          <i data-lucide="folders" class="w-10 h-10 mx-auto mb-2.5 opacity-40"></i>
          <p class="font-bold text-slate-700 text-sm">개설된 과목 바구니가 없습니다.</p>
          <p class="text-xs text-slate-400 mt-1">상단의 <strong>[+ 과목 바구니 생성]</strong>으로 새 과목을 추가해주세요.</p>
        </div>
      `;
      if (window.lucide) window.lucide.createIcons();
      return;
    }

    const masterMap = new Map(this.masterStudents.map(s => [s.id, s]));
    const regularPeriods = this.periods.filter(p => !p.isBreak);
    const activeDays = this.showWeekends ? this.days : this.days.filter(d => !d.isWeekend);

    let html = '';
    this.offeredSubjects.forEach(sub => {
      const color = this.getSubjectColor(sub);
      const studentIds = this.courseEnrollments[sub] || [];
      const schedules = this.courseSchedules[sub] || [];

      html += `
        <div class="bg-white rounded-xl border border-slate-200 shadow-xs flex flex-col overflow-hidden transition-all hover:shadow-md">
          
          <!-- 바구니 상단 헤더 -->
          <div class="p-3 border-b flex items-center justify-between"
               style="background-color: ${color.bg}; border-color: ${color.border};">
            <div class="flex items-center gap-2">
              <span class="w-3.5 h-3.5 rounded-full border shadow-2xs" 
                    style="background-color: ${color.text}; border-color: ${color.border};"></span>
              <h3 class="font-bold text-sm" style="color: ${color.text};">${sub}</h3>
              <span class="text-xs font-bold px-2 py-0.5 rounded-full bg-white/80 border text-slate-700 shadow-2xs">
                ${studentIds.length}명 편제
              </span>
            </div>
            <button type="button" 
                    onclick="window.app.handleDeleteSubjectBucket('${sub}')" 
                    class="text-slate-400 hover:text-red-600 p-1 transition" title="과목 바구니 삭제">
              <i data-lucide="trash-2" class="w-4 h-4"></i>
            </button>
          </div>

          <!-- 바구니 본문 -->
          <div class="p-3.5 flex-1 flex flex-col gap-3">
            
            <!-- 드롭존 (학생을 여기로 드래그 앤 드롭) -->
            <div class="bucket-drop-zone p-2.5 rounded-lg flex flex-col items-center justify-center gap-1 text-center cursor-pointer"
                 ondragover="window.app.handleBucketDragOver(event)"
                 ondragleave="window.app.handleBucketDragLeave(event)"
                 ondrop="window.app.handleBucketDrop(event, '${sub}')">
              <div class="flex items-center gap-1.5 text-blue-600 font-bold text-xs">
                <i data-lucide="user-plus" class="w-4 h-4"></i>
                <span>학생을 여기로 드래그하여 편제 (Drop)</span>
              </div>
              <span class="text-[10px] text-slate-400">좌측 학생 명단에서 끌어다 놓으세요</span>
            </div>

            <!-- 편제된 학생 태그/칩 목록 -->
            <div>
              <div class="flex items-center justify-between text-[11px] font-bold text-slate-700 mb-1.5">
                <span>편제된 수강생 목록 (${studentIds.length}명)</span>
                ${studentIds.length > 0 ? `
                  <button type="button" onclick="window.app.handleClearBucketStudents('${sub}')" class="text-[10px] text-slate-400 hover:text-red-500 font-normal">
                    전체 비우기
                  </button>
                ` : ''}
              </div>

              <div class="flex flex-wrap gap-1.5 max-h-32 overflow-y-auto p-1 bg-slate-50 rounded-lg border border-slate-100 min-h-[46px]">
                ${studentIds.length > 0 ? studentIds.map(stId => {
                  const m = masterMap.get(stId);
                  if (!m) return '';
                  const code = m.studentCode || ExcelService.getStudentCode(m);
                  return `
                    <span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold bg-white border border-slate-200 text-slate-800 shadow-2xs hover:border-blue-400 transition">
                      <span class="font-mono text-[10px] text-slate-500">${code}</span>
                      <span>${m.name}</span>
                      <button type="button" 
                              onclick="window.app.handleRemoveStudentFromCourse('${sub}', '${m.id}')" 
                              class="text-slate-400 hover:text-red-600 ml-0.5 font-bold" title="편제 제외">&times;</button>
                    </span>
                  `;
                }).join('') : `
                  <span class="text-xs text-slate-400 p-2 flex items-center justify-center w-full">편제된 학생이 없습니다.</span>
                `}
              </div>
            </div>

            <!-- 과목 시간표 시간 설정 바 (시간표 자동 배치) -->
            <div class="pt-2.5 border-t border-slate-100 flex flex-col gap-1.5">
              <div class="flex items-center justify-between text-[11px] font-bold text-slate-700">
                <span class="flex items-center gap-1">
                  <i data-lucide="clock" class="w-3.5 h-3.5 text-indigo-600"></i>
                  과목 시간 설정 (시간표 자동 배치)
                </span>
                <span class="text-[10px] text-indigo-600 font-semibold">${schedules.length}시간 편성</span>
              </div>

              <!-- 요일 + 교시 추가 폼 -->
              <div class="flex items-center gap-1">
                <select id="bucket-day-${sub}" class="flex-1 px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded font-medium text-slate-800">
                  ${activeDays.map(d => `<option value="${d.id}">${d.fullName}</option>`).join('')}
                </select>
                <select id="bucket-period-${sub}" class="flex-1 px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded font-medium text-slate-800">
                  ${regularPeriods.map(p => `<option value="${p.id}">${p.name} (${p.startTime})</option>`).join('')}
                </select>
                <button type="button" 
                        onclick="window.app.handleAddBucketSchedule('${sub}')" 
                        class="px-2.5 py-1 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded transition whitespace-nowrap shadow-2xs">
                  + 시간 추가
                </button>
              </div>

              <!-- 현재 편성된 시간 칩 목록 (여러 날, 여러 시간) -->
              <div class="flex flex-wrap gap-1 mt-1">
                ${schedules.length > 0 ? schedules.map(sch => {
                  const dayName = ExcelService.dayIdToName(sch.day);
                  const periodObj = this.periods.find(p => p.id === sch.periodId || (!sch.periodId && p.period === sch.period));
                  const periodName = periodObj ? periodObj.name : `${sch.period}교시`;
                  const sym = ExcelService.getStatusSymbol(sch.status || 'confirmed');
                  return `
                    <span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-bold bg-indigo-50 text-indigo-900 border border-indigo-200">
                      <span>${sym}</span>
                      <span>${dayName} ${periodName}</span>
                      <button type="button" 
                              onclick="window.app.handleRemoveBucketSchedule('${sub}', '${sch.id || (sch.day + '_' + sch.periodId)}')"
                              class="text-indigo-400 hover:text-red-600 ml-1" title="시간표에서 배치 해제">&times;</button>
                    </span>
                  `;
                }).join('') : `
                  <span class="text-[11px] text-slate-400">아직 시간표에 배치되지 않음</span>
                `}
              </div>

            </div>

          </div>

        </div>
      `;
    });

    container.innerHTML = html;
    if (window.lucide) window.lucide.createIcons();
  }

  handleBucketDragOver(e) {
    e.preventDefault();
    e.currentTarget.classList.add('drag-over');
  }

  handleBucketDragLeave(e) {
    e.currentTarget.classList.remove('drag-over');
  }

  handleBucketDrop(e, subjectName) {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-over');

    let data = this.dragData;
    if (!data) {
      try {
        const raw = e.dataTransfer.getData('text/plain');
        if (raw) data = JSON.parse(raw);
      } catch (err) {}
    }

    if (!data || data.type !== 'student') return;

    const studentIds = data.studentIds && data.studentIds.length > 0 ? data.studentIds : [data.studentId];
    if (!studentIds || studentIds.length === 0) return;

    if (!this.courseEnrollments[subjectName]) this.courseEnrollments[subjectName] = [];
    let addedCount = 0;

    studentIds.forEach(id => {
      if (!this.courseEnrollments[subjectName].includes(id)) {
        this.courseEnrollments[subjectName].push(id);
        addedCount++;
      }
    });

    this.rebuildStudentAssignments();
    this.saveState();
    this.render();

    if (addedCount > 0) {
      this.showToast(`${addedCount}명의 학생이 '${subjectName}' 과목에 편제되었습니다!`);
    } else {
      this.showToast(`이미 '${subjectName}' 과목에 편제된 학생입니다.`);
    }
  }

  handleRemoveStudentFromCourse(subjectName, studentId) {
    if (!this.courseEnrollments[subjectName]) return;
    this.courseEnrollments[subjectName] = this.courseEnrollments[subjectName].filter(id => id !== studentId);
    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`학생이 '${subjectName}' 과목에서 제외되었습니다.`);
  }

  handleClearBucketStudents(subjectName) {
    if (!confirm(`'${subjectName}' 과목의 모든 수강생 편제를 비우시겠습니까?`)) return;
    this.courseEnrollments[subjectName] = [];
    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`'${subjectName}' 과목의 편제 학생이 모두 비워졌습니다.`);
  }

  handleAddBucketSchedule(subjectName) {
    const daySelect = document.getElementById(`bucket-day-${subjectName}`);
    const periodSelect = document.getElementById(`bucket-period-${subjectName}`);
    if (!daySelect || !periodSelect) return;

    const dayId = daySelect.value;
    const periodId = periodSelect.value;
    const periodObj = this.periods.find(p => p.id === periodId);

    this.addCourseSchedule(subjectName, dayId, periodId, 'confirmed');
  }

  handleRemoveBucketSchedule(subjectName, scheduleId) {
    if (!this.courseSchedules[subjectName]) return;
    this.courseSchedules[subjectName] = this.courseSchedules[subjectName].filter(sch => {
      const schKey = sch.id || (sch.day + '_' + sch.periodId);
      return schKey !== scheduleId;
    });
    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`'${subjectName}' 과목의 시간표 배치가 해제되었습니다.`);
  }

  // ==========================================
  // [2단계: 탭 2] 과목 팔레트 (드래그)
  // ==========================================
  renderStep2Palette() {
    const container = document.getElementById('palette-subject-list');
    const badge = document.getElementById('step2-subject-count-badge');
    if (!container) return;

    if (badge) badge.textContent = `${this.offeredSubjects.length}개`;

    if (this.offeredSubjects.length === 0) {
      container.innerHTML = `
        <div class="py-12 text-center text-slate-400 text-xs">
          <p class="font-medium text-slate-600">등록된 과목이 없습니다.</p>
          <p class="text-[11px] text-slate-400 mt-1">[1단계 바구니 편제]에서 과목을 먼저 생성하세요.</p>
        </div>
      `;
      return;
    }

    let html = '';
    this.offeredSubjects.forEach(sub => {
      const color = this.getSubjectColor(sub);
      const studentIds = this.courseEnrollments[sub] || [];
      const schedules = this.courseSchedules[sub] || [];

      html += `
        <div class="draggable-palette-subject p-3 rounded-xl border transition-all hover:scale-[1.01] shadow-2xs"
             style="background-color: ${color.bg}; border-color: ${color.border};"
             draggable="true"
             ondragstart="window.app.handlePaletteDragStart(event, '${sub}')"
             ondragend="window.app.handlePaletteDragEnd(event)">
          
          <div class="flex items-center justify-between gap-1 mb-1.5">
            <div class="flex items-center gap-1.5 min-w-0">
              <i data-lucide="grip-vertical" class="w-3.5 h-3.5 shrink-0" style="color: ${color.text};"></i>
              <span class="font-bold text-xs truncate" style="color: ${color.text};">${sub}</span>
            </div>
            <span class="text-[11px] font-bold px-2 py-0.5 rounded-full bg-white/90 border text-slate-700 shadow-2xs">
              ${studentIds.length}명
            </span>
          </div>

          <div class="flex items-center justify-between text-[11px] text-slate-600">
            <span>
              ${schedules.length > 0 ? `
                <strong class="text-indigo-700">${schedules.length}시간</strong> 배치됨
              ` : `
                <span class="text-amber-600 font-semibold">미배치</span>
              `}
            </span>
            <span class="text-[10px] text-slate-500">시간표로 드래그 &rarr;</span>
          </div>

        </div>
      `;
    });

    container.innerHTML = html;
    if (window.lucide) window.lucide.createIcons();
  }

  handlePaletteDragStart(e, subjectName) {
    this.dragData = {
      type: 'palette-subject',
      subjectName: subjectName
    };
    e.dataTransfer.setData('text/plain', JSON.stringify(this.dragData));
    e.dataTransfer.effectAllowed = 'copy';
    e.currentTarget.classList.add('is-dragging');
  }

  handlePaletteDragEnd(e) {
    this.dragData = null;
    e.currentTarget.classList.remove('is-dragging');
    document.querySelectorAll('.timetable-cell.drag-over').forEach(c => c.classList.remove('drag-over'));
  }

  // ==========================================
  // [2단계: 탭 2] 종합 시간표 그리드 & 드래그(이동)/복사(Ctrl)
  // ==========================================
  renderTimetable() {
    const thead = document.getElementById('timetable-thead');
    const tbody = document.getElementById('timetable-tbody');
    if (!thead || !tbody) return;

    this.sortPeriods();
    const activeDays = this.showWeekends ? this.days : this.days.filter(d => !d.isWeekend);

    // 1. 헤더 (요일)
    let theadHtml = `
      <tr class="bg-slate-100/90 border-b border-slate-200">
        <th class="py-2.5 px-3 text-center font-bold text-slate-700 w-28 shrink-0">교시 / 시간</th>
    `;
    activeDays.forEach(day => {
      const isWeekend = day.isWeekend;
      theadHtml += `
        <th class="py-2.5 px-3 text-center font-bold text-xs ${
          isWeekend ? 'bg-red-50 text-red-800' : 'text-slate-800'
        } border-l border-slate-200 min-w-[170px]">
          ${day.fullName} ${isWeekend ? '<span class="text-[10px] text-red-500">(주말)</span>' : ''}
        </th>
      `;
    });
    theadHtml += `</tr>`;
    thead.innerHTML = theadHtml;

    // 2. 바디 (교시 행)
    let bodyHtml = '';
    const masterMap = new Map(this.masterStudents.map(s => [s.id, s]));

    this.periods.forEach(period => {
      if (period.isBreak) {
        // 점심시간/휴식 행
        bodyHtml += `
          <tr class="bg-slate-50/80 border-b border-slate-200 text-slate-500">
            <td class="py-2 px-3 text-center font-semibold text-xs border-r border-slate-200 bg-slate-100/60">
              ☕ ${period.name}
              <div class="text-[10px] text-slate-400 font-mono">${period.startTime}~${period.endTime}</div>
            </td>
            <td colspan="${activeDays.length}" class="py-2 px-4 text-center font-medium text-xs text-slate-400 italic">
              ${period.name} (${period.startTime} ~ ${period.endTime})
            </td>
          </tr>
        `;
      } else {
        // 정규 수업 교시 행
        bodyHtml += `
          <tr class="border-b border-slate-200 hover:bg-slate-50/30 transition-colors">
            <!-- 교시 헤더 -->
            <td class="py-2 px-2.5 text-center text-xs font-bold text-slate-700 border-r border-slate-200 bg-slate-50/60 align-middle">
              <div class="font-extrabold text-slate-900">${period.name}</div>
              <div class="text-[10px] text-slate-500 font-mono mt-0.5">${period.startTime}~${period.endTime}</div>
              <button type="button" 
                      onclick="window.app.handleDeletePeriod('${period.id}')"
                      class="text-[10px] text-slate-300 hover:text-red-500 mt-1 transition" title="교시 삭제">
                삭제
              </button>
            </td>
        `;

        // 각 요일 셀
        activeDays.forEach(day => {
          bodyHtml += `
            <td class="timetable-cell p-1.5 border-l border-slate-200 align-top ${day.isWeekend ? 'bg-red-50/20' : ''}"
                data-day="${day.id}"
                data-period-id="${period.id}"
                ondragover="window.app.handleTimetableCellDragOver(event)"
                ondragleave="window.app.handleTimetableCellDragLeave(event)"
                ondrop="window.app.handleTimetableCellDrop(event, '${day.id}', '${period.id}')">
              
              <div class="min-h-[72px] flex flex-col gap-1.5">
          `;

          // 이 셀에 배치된 과목들 검색
          const scheduledEntries = [];
          Object.entries(this.courseSchedules).forEach(([subject, schedules]) => {
            schedules.forEach(sch => {
              if (sch.day === day.id && (sch.periodId === period.id || (!sch.periodId && sch.period === period.period))) {
                scheduledEntries.push({
                  subject,
                  schedule: sch
                });
              }
            });
          });

          if (scheduledEntries.length > 0) {
            scheduledEntries.forEach(({ subject, schedule }) => {
              const color = this.getSubjectColor(subject);
              const enrolledIds = this.courseEnrollments[subject] || [];
              const status = schedule.status || 'confirmed';
              const symHtml = ExcelService.getStatusHtml(status);
              const isCanceled = status === 'canceled';
              const isAdjust = status === 'adjust';

              // 학생 텍스트 요약
              let rosterSummary = '편제 학생 없음';
              if (enrolledIds.length > 0) {
                const firstM = masterMap.get(enrolledIds[0]);
                const firstName = firstM ? firstM.name : '';
                if (enrolledIds.length === 1) {
                  rosterSummary = `${firstName}`;
                } else {
                  rosterSummary = `${firstName} 외 ${enrolledIds.length - 1}명`;
                }
              }

              bodyHtml += `
                <div class="subject-card p-2 rounded-lg border text-xs transition-all relative ${
                  isCanceled ? 'is-canceled' : ''
                } ${isAdjust ? 'is-adjust' : ''}"
                     style="background-color: ${isCanceled ? '#E2E8F0' : color.bg}; border-color: ${isCanceled ? '#CBD5E1' : color.border};"
                     draggable="true"
                     ondragstart="window.app.handleTimetableCardDragStart(event, '${subject}', '${day.id}', '${period.id}')"
                     ondragend="window.app.handleTimetableCardDragEnd(event)">
                  
                  <!-- 카드 상단: 상태 기호 + 과목명 + 삭제 버튼 -->
                  <div class="flex items-center justify-between gap-1 mb-1">
                    <div class="flex items-center gap-1 min-w-0">
                      <!-- 상태 토글 버튼 (확정 -> 조정 -> 폐강) -->
                      <button type="button" 
                              onclick="event.stopPropagation(); window.app.toggleCourseScheduleStatus('${subject}', '${schedule.id || (day.id + '_' + period.id)}')"
                              class="hover:scale-110 transition shrink-0" title="클릭하여 상태 전환 (🟢확정 -> ▲조정 -> ✖️폐강)">
                        ${symHtml}
                      </button>
                      <span class="font-extrabold truncate" style="color: ${isCanceled ? '#64748B' : color.text};">
                        ${subject}
                      </span>
                    </div>

                    <!-- 이 시간표 셀에서만 삭제 -->
                    <button type="button" 
                            onclick="event.stopPropagation(); window.app.handleRemoveCourseScheduleSlot('${subject}', '${schedule.id || (day.id + '_' + period.id)}')"
                            class="text-slate-400 hover:text-red-600 font-bold px-1 transition" title="이 시간 배치 취소">
                      &times;
                    </button>
                  </div>

                  <!-- 카드 하단: 학생 명단 칩 & 명단 팝업 버튼 -->
                  <div class="flex items-center justify-between text-[11px]">
                    <button type="button" 
                            onclick="event.stopPropagation(); window.app.openCourseRosterModal('${subject}', '${day.id}', '${period.id}')"
                            class="text-slate-700 hover:text-blue-700 font-semibold truncate hover:underline flex items-center gap-1"
                            title="전체 수강 학생 명단 보기">
                      <i data-lucide="users" class="w-3 h-3 text-slate-500"></i>
                      <span>${rosterSummary}</span>
                    </button>
                    <span class="text-[10px] font-bold text-slate-500 shrink-0">
                      (${enrolledIds.length}명)
                    </span>
                  </div>

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

  // ==========================================
  // 시간표 카드 드래그 & 셀 드롭 (이동 & Ctrl 복사)
  // ==========================================
  handleTimetableCardDragStart(e, subjectName, fromDay, fromPeriodId) {
    const isCopy = e.ctrlKey || e.metaKey || this.isCtrlPressed;
    this.dragData = {
      type: 'timetable-subject',
      subjectName: subjectName,
      fromDay: fromDay,
      fromPeriodId: fromPeriodId,
      isCtrlCopy: isCopy
    };

    e.dataTransfer.setData('text/plain', JSON.stringify(this.dragData));
    e.dataTransfer.effectAllowed = isCopy ? 'copy' : 'copyMove';
    e.currentTarget.classList.add('is-dragging');

    if (isCopy) {
      const banner = document.getElementById('ctrl-copy-indicator');
      if (banner) banner.classList.remove('hidden');
    }
  }

  handleTimetableCardDragEnd(e) {
    this.dragData = null;
    e.currentTarget.classList.remove('is-dragging');
    document.querySelectorAll('.timetable-cell.drag-over').forEach(c => c.classList.remove('drag-over'));
  }

  handleTimetableCellDragOver(e) {
    e.preventDefault();
    const isCopy = e.ctrlKey || e.metaKey || this.isCtrlPressed || (this.dragData && this.dragData.isCtrlCopy);
    e.dataTransfer.dropEffect = isCopy ? 'copy' : 'move';
    const cell = e.currentTarget;
    if (!cell.classList.contains('drag-over')) {
      cell.classList.add('drag-over');
    }
  }

  handleTimetableCellDragLeave(e) {
    e.currentTarget.classList.remove('drag-over');
  }

  handleTimetableCellDrop(e, targetDay, targetPeriodId) {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-over');

    let data = this.dragData;
    if (!data) {
      try {
        const raw = e.dataTransfer.getData('text/plain');
        if (raw) data = JSON.parse(raw);
      } catch (err) {}
    }

    if (!data) return;

    const isCopyMode = e.ctrlKey || e.metaKey || this.isCtrlPressed || data.isCtrlCopy;
    const targetPeriod = this.periods.find(p => p.id === targetPeriodId);
    const dayName = ExcelService.dayIdToName(targetDay);
    const periodName = targetPeriod ? targetPeriod.name : '';

    // 1. 좌측 과목 팔레트에서 새 과목을 끌고 온 경우
    if (data.type === 'palette-subject') {
      this.addCourseSchedule(data.subjectName, targetDay, targetPeriodId, 'confirmed');
      const studentCount = (this.courseEnrollments[data.subjectName] || []).length;
      this.showToast(`'${data.subjectName}' 과목 (학생 ${studentCount}명)이 ${dayName} ${periodName}에 배치되었습니다.`);
      return;
    }

    // 2. 시간표 내의 과목 카드를 다른 셀로 끌고 온 경우
    if (data.type === 'timetable-subject') {
      const { subjectName, fromDay, fromPeriodId } = data;

      // 같은 위치로 드롭한 경우 무시
      if (fromDay === targetDay && fromPeriodId === targetPeriodId) {
        return;
      }

      if (isCopyMode) {
        // [복사 (Copy)]: 기존 슬롯은 유지하고 새 슬롯 추가!
        this.addCourseSchedule(subjectName, targetDay, targetPeriodId, 'confirmed');
        this.showToast(`[복사 완료] '${subjectName}' 과목이 ${dayName} ${periodName}에 추가 배치되었습니다.`);
      } else {
        // [이동 (Move)]: 기존 위치 슬롯을 새 위치로 변경!
        this.moveCourseSchedule(subjectName, fromDay, fromPeriodId, targetDay, targetPeriodId);
        this.showToast(`'${subjectName}' 과목이 ${dayName} ${periodName}으로 이동되었습니다.`);
      }
    }
  }

  addCourseSchedule(subjectName, day, periodId, status = 'confirmed') {
    if (!this.courseSchedules[subjectName]) this.courseSchedules[subjectName] = [];
    const periodObj = this.periods.find(p => p.id === periodId);

    // 중복 체크
    const exists = this.courseSchedules[subjectName].some(sch => sch.day === day && sch.periodId === periodId);
    if (!exists) {
      this.courseSchedules[subjectName].push({
        id: `sch_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
        day: day,
        periodId: periodId,
        period: periodObj ? periodObj.period : null,
        status: status
      });
      this.rebuildStudentAssignments();
      this.saveState();
      this.render();
    }
  }

  moveCourseSchedule(subjectName, fromDay, fromPeriodId, toDay, toPeriodId) {
    if (!this.courseSchedules[subjectName]) return;
    const periodObj = this.periods.find(p => p.id === toPeriodId);

    const slot = this.courseSchedules[subjectName].find(sch => sch.day === fromDay && sch.periodId === fromPeriodId);
    if (slot) {
      slot.day = toDay;
      slot.periodId = toPeriodId;
      slot.period = periodObj ? periodObj.period : null;
      this.rebuildStudentAssignments();
      this.saveState();
      this.render();
    }
  }

  handleRemoveCourseScheduleSlot(subjectName, scheduleId) {
    if (!this.courseSchedules[subjectName]) return;
    this.courseSchedules[subjectName] = this.courseSchedules[subjectName].filter(sch => {
      const schKey = sch.id || (sch.day + '_' + sch.periodId);
      return schKey !== scheduleId;
    });
    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`'${subjectName}' 배치가 취소되었습니다.`);
  }

  toggleCourseScheduleStatus(subjectName, scheduleId) {
    if (!this.courseSchedules[subjectName]) return;
    const slot = this.courseSchedules[subjectName].find(sch => {
      const schKey = sch.id || (sch.day + '_' + sch.periodId);
      return schKey === scheduleId;
    });
    if (!slot) return;

    if (slot.status === 'confirmed') slot.status = 'adjust';
    else if (slot.status === 'adjust') slot.status = 'canceled';
    else slot.status = 'confirmed';

    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`'${subjectName}' 상태가 [${ExcelService.getStatusText(slot.status)}]로 변경되었습니다.`);
  }

  // ==========================================
  // 과목 수강생 출석부 / 명단 모달
  // ==========================================
  openCourseRosterModal(subjectName, dayId, periodId) {
    const modal = document.getElementById('course-roster-modal');
    const title = document.getElementById('roster-modal-title');
    const subtitle = document.getElementById('roster-modal-subtitle');
    const list = document.getElementById('roster-modal-list');
    const countBadge = document.getElementById('roster-modal-count');
    if (!modal) return;

    const dayName = ExcelService.dayIdToName(dayId);
    const periodObj = this.periods.find(p => p.id === periodId);
    const periodName = periodObj ? periodObj.name : '';

    title.innerHTML = `
      <i data-lucide="users" class="w-4 h-4 text-blue-600"></i>
      <span>'${subjectName}' 수강 학생 명단</span>
    `;
    subtitle.textContent = `${dayName}요일 ${periodName} 수업 | 상지대관령고등학교`;

    const enrolledIds = this.courseEnrollments[subjectName] || [];
    const masterMap = new Map(this.masterStudents.map(s => [s.id, s]));

    if (countBadge) countBadge.textContent = `총 ${enrolledIds.length}명 수강`;

    if (enrolledIds.length === 0) {
      list.innerHTML = `<p class="py-8 text-center text-slate-400 text-xs">수강생이 없습니다.</p>`;
    } else {
      list.innerHTML = enrolledIds.map((id, idx) => {
        const m = masterMap.get(id);
        if (!m) return '';
        const code = m.studentCode || ExcelService.getStudentCode(m);
        return `
          <div class="py-2.5 px-3 flex items-center justify-between text-xs hover:bg-slate-50">
            <div class="flex items-center gap-3">
              <span class="w-5 text-center text-slate-400 font-mono text-[11px]">${idx + 1}</span>
              <span class="font-mono font-bold text-slate-900">${code}</span>
              <span class="font-bold text-slate-800">${m.name}</span>
            </div>
            <button type="button" 
                    onclick="window.app.handleRemoveStudentFromCourse('${subjectName}', '${m.id}'); window.app.openCourseRosterModal('${subjectName}', '${dayId}', '${periodId}');" 
                    class="text-xs text-slate-400 hover:text-red-600 font-medium transition" title="수강 제외">
              제외
            </button>
          </div>
        `;
      }).join('');
    }

    modal.classList.remove('hidden');
    modal.classList.add('flex');
    if (window.lucide) window.lucide.createIcons();
  }

  closeCourseRosterModal() {
    const modal = document.getElementById('course-roster-modal');
    if (!modal) return;
    modal.classList.add('hidden');
    modal.classList.remove('flex');
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
    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.closeAddPeriodModal();
    this.showToast(`'${name}' 교시가 추가되었습니다.`);
  }

  handleDeletePeriod(periodId) {
    const period = this.periods.find(p => p.id === periodId);
    if (!period) return;

    if (!confirm(`'${period.name}' 교시를 삭제하시겠습니까? 해당 교시에 배치된 과목은 배정이 해제됩니다.`)) {
      return;
    }

    // 해당 교시의 스케줄 삭제
    Object.keys(this.courseSchedules).forEach(sub => {
      this.courseSchedules[sub] = this.courseSchedules[sub].filter(sch => sch.periodId !== periodId);
    });

    this.periods = this.periods.filter(p => p.id !== periodId);
    this.sortPeriods();
    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`${period.name}가 삭제되었습니다.`);
  }

  // ==========================================
  // 상세 학생 등록 모달
  // ==========================================
  openStudentModal() {
    const modal = document.getElementById('student-modal');
    if (!modal) return;
    document.getElementById('form-add-student').reset();
    const dl = document.getElementById('modal-subject-list');
    if (dl) {
      dl.innerHTML = this.offeredSubjects.map(s => `<option value="${s}"></option>`).join('');
    }
    modal.classList.remove('hidden');
    modal.classList.add('flex');
  }

  closeStudentModal() {
    const modal = document.getElementById('student-modal');
    if (!modal) return;
    modal.classList.add('hidden');
    modal.classList.remove('flex');
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
