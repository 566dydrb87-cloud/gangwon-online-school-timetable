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
    this.storageKey = 'gw_online_school_timetable_data_v5';
    this.legacyStorageKey = 'gw_online_school_timetable_data_v4';
    this.legacyStorageKeyV3 = 'gw_online_school_timetable_data_v3';
    this.activeYear = '2027'; // 기본 활성 학년도: 2027학년도
    this.activeGrade = '2'; // 기본 활성 학년: 2학년 ('2' | '3')
    this.activeSemester = '1'; // 기본 활성 학기: 1학기 ('1' | '2')
    this.academicYears = ['2027', '2026'];
    this.grades = ['2', '3'];
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
    this.studentSortKey = 'code'; // 'code' | 'name' | 'subject'
    this.excelPoolStorageKey = 'gw_excel_student_pool_directory_v1';
    this.excelStudentPool = []; // 업로드된 전체 학생 명부 풀 (검색 및 선택 등록용)
    this.isCtrlPressed = false;
    this.dragData = null; // { type, studentId, studentIds, subjectName, fromDay, fromPeriodId }
    this.undoStack = []; // 실행 취소(되돌리기) 스냅샷 스택
    this.redoStack = []; // 다시 실행(재실행) 스냅샷 스택
    this.maxUndoDepth = 15;
    this.highlightedSubject = null; // 과목 팔레트에서 선택하여 시간표에 하이라이트할 과목명

    this.init();
  }

  async init() {
    this.loadState();
    this.loadExcelStudentPool();
    this.sortPeriods();
    this.sortOfferedSubjects();
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
      courseStatus: {},
      students: [],
      offeredSubjects: baseSubjects ? [...baseSubjects] : [...DEFAULT_OFFERED_SUBJECTS],
      subjectColorMap: baseColorMap ? JSON.parse(JSON.stringify(baseColorMap)) : {},
      showWeekends: true
    };
  }

  ensureYearSemesterData(year = this.activeYear, grade = this.activeGrade || '2', semester = this.activeSemester || '1') {
    if (!this.yearsData[year]) {
      this.yearsData[year] = {};
    }

    // 구 버전 데이터 마이그레이션 (yearsData[year]['1'] 바로 아래에 periods가 있는 경우)
    const isOldStructure = (this.yearsData[year]['1'] && this.yearsData[year]['1'].periods) ||
                           (this.yearsData[year]['2'] && this.yearsData[year]['2'].periods) ||
                           this.yearsData[year].periods;
    if (isOldStructure) {
      const oldSem1 = this.yearsData[year]['1'] || (this.yearsData[year].periods ? this.yearsData[year] : null) || this.createDefaultSemesterData();
      const oldSem2 = this.yearsData[year]['2'] || this.createDefaultSemesterData();
      this.yearsData[year] = {
        '2': {
          '1': oldSem1,
          '2': oldSem2
        },
        '3': {
          '1': this.createDefaultSemesterData(),
          '2': this.createDefaultSemesterData()
        }
      };
    }

    // 2학년, 3학년 컨테이너 확보
    if (!this.yearsData[year]['2']) {
      this.yearsData[year]['2'] = {
        '1': this.createDefaultSemesterData(),
        '2': this.createDefaultSemesterData()
      };
    }
    if (!this.yearsData[year]['3']) {
      this.yearsData[year]['3'] = {
        '1': this.createDefaultSemesterData(),
        '2': this.createDefaultSemesterData()
      };
    }

    if (!this.yearsData[year][grade]) {
      this.yearsData[year][grade] = {
        '1': this.createDefaultSemesterData(),
        '2': this.createDefaultSemesterData()
      };
    }
    if (!this.yearsData[year][grade]['1']) {
      this.yearsData[year][grade]['1'] = this.createDefaultSemesterData();
    }
    if (!this.yearsData[year][grade]['2']) {
      this.yearsData[year][grade]['2'] = this.createDefaultSemesterData();
    }

    const d = this.yearsData[year][grade][semester];
    if (!d.periods || !Array.isArray(d.periods) || d.periods.length === 0) {
      d.periods = JSON.parse(JSON.stringify(DEFAULT_PERIODS));
    }
    if (!d.days || !Array.isArray(d.days) || d.days.length === 0) {
      d.days = JSON.parse(JSON.stringify(DEFAULT_DAYS));
    }
    if (!d.offeredSubjects) {
      d.offeredSubjects = [];
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
              grade: Number(grade) || s.grade || 2,
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
    this.ensureYearSemesterData(this.activeYear, this.activeGrade, this.activeSemester);
    const d = this.yearsData[this.activeYear][this.activeGrade][this.activeSemester];
    this.periods = d.periods;
    this.days = d.days;
    this.masterStudents = d.masterStudents || [];
    this.courseEnrollments = d.courseEnrollments || {};
    this.courseSchedules = d.courseSchedules || {};
    this.courseStatus = d.courseStatus || {};
    this.offeredSubjects = d.offeredSubjects || [];
    this.sortOfferedSubjects();
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
            grade: Number(this.activeGrade) || s.grade || 2,
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
    this.sortOfferedSubjects();
    this.rebuildStudentAssignments();
    this.ensureYearSemesterData(this.activeYear, this.activeGrade, this.activeSemester);
    this.yearsData[this.activeYear][this.activeGrade][this.activeSemester] = {
      periods: this.periods,
      days: this.days,
      masterStudents: this.masterStudents,
      courseEnrollments: this.courseEnrollments,
      courseSchedules: this.courseSchedules,
      courseStatus: this.courseStatus || {},
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
        // v4, v3 레거시 키 순차 확인
        saved = localStorage.getItem(this.legacyStorageKey);
      }
      if (!saved) {
        saved = localStorage.getItem(this.legacyStorageKeyV3);
      }

      if (saved) {
        const data = JSON.parse(saved);
        this.yearsData = data.yearsData || data.years || {};
        this.academicYears = data.academicYears || Object.keys(this.yearsData);
        if (!this.academicYears.includes('2027')) this.academicYears.unshift('2027');
        if (!this.academicYears.includes('2026')) this.academicYears.push('2026');

        // 기본 작업 학년도: 2027학년도, 기본 학년: 2학년
        this.activeYear = data.activeYear || '2027';
        this.activeGrade = data.activeGrade || '2';
        this.activeSemester = data.activeSemester || '1';

        this.ensureYearSemesterData('2027', '2', '1');
        this.ensureYearSemesterData('2027', '2', '2');
        this.ensureYearSemesterData('2027', '3', '1');
        this.ensureYearSemesterData('2027', '3', '2');
        this.ensureYearSemesterData('2026', '2', '1');
        this.ensureYearSemesterData('2026', '2', '2');

        this.loadActiveYearSemesterData();
        return;
      }
    } catch (e) {
      console.error('Failed to load state:', e);
    }

    // 초기 상태 (2027학년도 2학년 1학기 기본)
    this.activeYear = '2027';
    this.activeGrade = '2';
    this.activeSemester = '1';
    this.academicYears = ['2027', '2026'];
    this.ensureYearSemesterData('2027', '2', '1');
    this.ensureYearSemesterData('2027', '3', '1');
    this.loadActiveYearSemesterData();
  }

  saveState() {
    try {
      this.saveActiveYearSemesterData();
      const data = {
        activeYear: this.activeYear,
        activeGrade: this.activeGrade,
        activeSemester: this.activeSemester,
        academicYears: this.academicYears,
        grades: this.grades,
        yearsData: this.yearsData,
        periods: this.periods,
        days: this.days,
        students: this.students,
        masterStudents: this.masterStudents,
        courseEnrollments: this.courseEnrollments,
        courseSchedules: this.courseSchedules,
        courseStatus: this.courseStatus || {},
        offeredSubjects: this.offeredSubjects,
        subjectColorMap: this.subjectColorMap,
        showWeekends: this.showWeekends,
        lastUpdated: new Date().toISOString()
      };
      localStorage.setItem(this.storageKey, JSON.stringify(data));
      this.updateStats();
      this.renderAcademicYearTabs();
      this.renderAcademicGradeTabs();

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
      if (data.activeGrade && ['2', '3'].includes(String(data.activeGrade))) this.activeGrade = String(data.activeGrade);
      if (data.activeSemester) this.activeSemester = data.activeSemester;
    }

    this.ensureYearSemesterData(this.activeYear, this.activeGrade, this.activeSemester);
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
    this.showToast(`[${this.activeYear}학년도 ${this.activeGrade}학년 ${this.activeSemester}학기] 시간표로 전환되었습니다.`);
  }

  // 대상 학년 전환 (2학년 / 3학년)
  switchGrade(grade) {
    if (this.activeGrade === grade) return;
    this.saveActiveYearSemesterData();
    this.activeGrade = grade;
    this.loadActiveYearSemesterData();
    this.sortPeriods();
    this.buildSubjectColorMap();
    this.saveState();
    this.render();
    this.showToast(`[${this.activeYear}학년도 ${this.activeGrade}학년 ${this.activeSemester}학기] 작업 환경으로 전환되었습니다.`);
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

    this.pushUndoSnapshot(`'${yearToDelete}학년도' 삭제`);

    delete this.yearsData[yearToDelete];
    this.academicYears = this.academicYears.filter(y => y !== yearToDelete);
    if (this.activeYear === yearToDelete) {
      this.activeYear = this.academicYears[0];
    }
    this.loadActiveYearSemesterData();
    this.saveState();
    this.render();
    this.showToast(`'${yearToDelete}학년도'가 삭제되었습니다.`, true);
  }

  // ==========================================
  // 일과표 교시 정렬
  // ==========================================
  sortPeriods() {
    this.periods.sort((a, b) => this.comparePeriods(a, b));
  }

  // ==========================================
  // 과목 목록 정렬 (과목명 이름 순 ㄱ부터 가나다순 정렬)
  // ==========================================
  sortOfferedSubjects() {
    if (!this.offeredSubjects || !Array.isArray(this.offeredSubjects)) {
      this.offeredSubjects = [];
      return;
    }
    this.offeredSubjects.sort((a, b) => (a || '').localeCompare(b || '', 'ko'));
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
    // 1. 키보드 Ctrl/Cmd 키 추적 (복사 모드 토글) 및 Ctrl+Z 실행 취소
    window.addEventListener('keydown', (e) => {
      // Ctrl + Z 실행 취소 / Ctrl + Y 또는 Ctrl + Shift + Z 다시 실행
      if (e.ctrlKey || e.metaKey) {
        const isZ = e.key === 'z' || e.key === 'Z';
        const isY = e.key === 'y' || e.key === 'Y';
        if (isZ || isY) {
          const activeTag = document.activeElement ? document.activeElement.tagName.toLowerCase() : '';
          const isInput = activeTag === 'input' || activeTag === 'textarea' || document.activeElement.isContentEditable;
          if (!isInput) {
            if (isZ && !e.shiftKey) {
              e.preventDefault();
              this.handleUndo();
              return;
            } else if (isY || (isZ && e.shiftKey)) {
              e.preventDefault();
              this.handleRedo();
              return;
            }
          }
        }
      }

      // ESC 키 입력 시 과목 하이라이트 해제
      if (e.key === 'Escape' && this.highlightedSubject) {
        this.clearSubjectHighlight();
        return;
      }

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

    // 창 전환 시 Ctrl 상태 고착 방지
    window.addEventListener('blur', () => {
      this.isCtrlPressed = false;
      const banner = document.getElementById('ctrl-copy-indicator');
      if (banner) banner.classList.add('hidden');
      document.body.classList.remove('is-copy-mode');
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

    // 6. 기준 학생명부 및 통합 엑셀 업로드 이벤트 분리
    document.getElementById('step1-file-upload-master-roster')?.addEventListener('change', async (e) => {
      const file = e.target.files?.[0];
      if (file) {
        await this.handleMasterRosterUpload(file);
      }
      e.target.value = '';
    });

    document.getElementById('step1-file-upload-unified-excel')?.addEventListener('change', async (e) => {
      if (e.target.files && e.target.files.length > 0) {
        await this.handleUnifiedExcelUpload(e.target.files);
      }
      e.target.value = '';
    });

    // 기존 호환성 유지
    document.getElementById('file-upload-student')?.addEventListener('change', async (e) => {
      if (e.target.files && e.target.files.length > 0) {
        await this.handleUnifiedExcelUpload(e.target.files);
      }
      e.target.value = '';
    });
    document.getElementById('step1-file-upload-student')?.addEventListener('change', async (e) => {
      if (e.target.files && e.target.files.length > 0) {
        await this.handleUnifiedExcelUpload(e.target.files);
      }
      e.target.value = '';
    });

    // 6-1. 엑셀 풀 실시간 학생명 검색 & 엔터 등록
    const inputSearchPool = document.getElementById('input-search-student-pool');
    const btnClearSearch = document.getElementById('btn-clear-pool-search');

    inputSearchPool?.addEventListener('input', (e) => {
      this.handleSearchPoolInput(e.target.value);
    });

    inputSearchPool?.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        this.handleSearchPoolEnter(e.target.value);
      }
    });

    btnClearSearch?.addEventListener('click', () => {
      this.clearPoolSearch();
      inputSearchPool?.focus();
    });

    // 바깥 클릭 시 검색 추천창 닫기
    document.addEventListener('click', (e) => {
      const dropdown = document.getElementById('pool-search-results-dropdown');
      if (dropdown && !dropdown.contains(e.target) && e.target !== inputSearchPool) {
        dropdown.classList.add('hidden');
      }
    });

    // 6-2. 학년도/학기 전송 모달 라디오 변경 시 과목 선택 드롭다운 토글
    document.querySelectorAll('input[name="transfer-scope"]').forEach(radio => {
      radio.addEventListener('change', (e) => {
        const sel = document.getElementById('transfer-subject-select');
        if (sel) {
          sel.disabled = e.target.value !== 'subject';
        }
      });
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
      this.handleDownloadUnifiedExcel();
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
  // 통합 엑셀 받기 (해당 학기뿐 아니라 1·2학기 전체 데이터를 포함하여 다운로드)
  // ==========================================
  handleDownloadUnifiedExcel() {
    // 1. 현재 작업 중인 활성 학기 데이터를 먼저 저장
    this.saveActiveYearSemesterData();

    const targetYear = this.activeYear || '2026';
    this.ensureYearSemesterData(targetYear, '1');
    this.ensureYearSemesterData(targetYear, '2');

    const sem1Data = this.yearsData[targetYear]['1'];
    const sem2Data = this.yearsData[targetYear]['2'];

    const sem1Students = this.buildFlatStudentListFromSemesterData(sem1Data);
    const sem2Students = this.buildFlatStudentListFromSemesterData(sem2Data);

    const totalStudentsCount = sem1Students.length + sem2Students.length;

    if (totalStudentsCount > 0) {
      ExcelService.exportYearFullDataExcel(targetYear, [
        { semester: '1', students: sem1Students, periods: sem1Data.periods || this.periods },
        { semester: '2', students: sem2Students, periods: sem2Data.periods || this.periods }
      ]);
      this.showToast(`[${targetYear}학년도 전체(1·2학기)] 시간표 통합 엑셀이 다운로드되었습니다. (1학기 ${sem1Students.length}건, 2학기 ${sem2Students.length}건)`, true);
    } else {
      // 1·2학기 모두 작성된 데이터가 없는 경우 -> 1학기 10행, 2학기 10행의 표준 빈 양식 다운로드
      ExcelService.downloadStudentTemplate(targetYear);
      this.showToast(`[${targetYear}학년도 1·2학기] 시간표 통합 표준 빈 양식이 다운로드되었습니다.`, true);
    }
  }

  // ==========================================
  // 1단계 좌측 패널: 학생 명단 및 과목 편제 현황 엑셀 다운로드 (1학기 + 2학기 연속 출력)
  // ==========================================
  exportStep1StudentRosterExcel() {
    // 1. 현재 작업 중인 활성 학기 데이터를 확실히 저장 동기화
    this.saveActiveYearSemesterData();

    const targetYear = this.activeYear || '2026';
    const targetGrade = this.activeGrade || '2';

    this.ensureYearSemesterData(targetYear, targetGrade, '1');
    this.ensureYearSemesterData(targetYear, targetGrade, '2');

    const sem1Data = (this.yearsData[targetYear] && this.yearsData[targetYear][targetGrade] && this.yearsData[targetYear][targetGrade]['1']) || {};
    const sem2Data = (this.yearsData[targetYear] && this.yearsData[targetYear][targetGrade] && this.yearsData[targetYear][targetGrade]['2']) || {};

    const sem1Enrollments = sem1Data.courseEnrollments || {};
    const sem2Enrollments = sem2Data.courseEnrollments || {};

    // 학생 정렬 헬퍼 (미편성 최상단 + 학번순)
    const sortRosterStudents = (students, enrollments) => {
      if (!students || !Array.isArray(students)) return [];
      const list = [...students];
      list.sort((a, b) => {
        let countA = 0;
        let countB = 0;
        Object.values(enrollments || {}).forEach(ids => {
          if (ids && ids.includes(a.id)) countA++;
          if (ids && ids.includes(b.id)) countB++;
        });
        const unassignedA = countA === 0 ? 0 : 1;
        const unassignedB = countB === 0 ? 0 : 1;
        if (unassignedA !== unassignedB) {
          return unassignedA - unassignedB; // 미편성(0) 학생 우선 배치
        }
        const codeA = a.studentCode || ExcelService.getStudentCode(a);
        const codeB = b.studentCode || ExcelService.getStudentCode(b);
        return codeA.localeCompare(codeB, undefined, { numeric: true });
      });
      return list;
    };

    const sortedSem1Students = sortRosterStudents(sem1Data.masterStudents || [], sem1Enrollments);
    const sortedSem2Students = sortRosterStudents(sem2Data.masterStudents || [], sem2Enrollments);

    const totalCount = sortedSem1Students.length + sortedSem2Students.length;
    if (totalCount === 0) {
      this.showToast('등록된 학생 명단이 없습니다. 학생 명부를 먼저 등록해주세요.');
      return;
    }

    ExcelService.exportStudentRosterExcel({
      academicYear: targetYear,
      grade: targetGrade,
      schoolName: '강원온라인학교',
      sem1Students: sortedSem1Students,
      sem1Enrollments: sem1Enrollments,
      sem2Students: sortedSem2Students,
      sem2Enrollments: sem2Enrollments
    });

    this.showToast(`[${targetYear}학년도 ${targetGrade}학년] 1·2학기 통합 명단(1학기 ${sortedSem1Students.length}명, 2학기 ${sortedSem2Students.length}명) 엑셀이 다운로드되었습니다.`, true);
  }

  // ==========================================
  // 학기·과목별 수강자 명단 수집 헬퍼 (1학기 -> 2학기, 학년/반/번호 순 정렬)
  // ==========================================
  collectSubjectRosterData(targetYear = this.activeYear || '2026') {
    this.saveActiveYearSemesterData();

    const yearData = this.yearsData[targetYear] || {};
    const gradeKeys = Object.keys(yearData);

    // 학생 정렬: 1학년 먼저, 다음 2학년 등 (학년 -> 반 -> 번호 -> 이름순)
    const compareStudents = (a, b) => {
      const gA = Number(a.grade) || 999;
      const gB = Number(b.grade) || 999;
      if (gA !== gB) return gA - gB; // 1학년 먼저, 2학년, 3학년 순

      const cA = Number(a.classNum) || 999;
      const cB = Number(b.classNum) || 999;
      if (cA !== cB) return cA - cB; // 반 오름차순

      const nA = Number(a.studentNum) || 999;
      const nB = Number(b.studentNum) || 999;
      if (nA !== nB) return nA - nB; // 번호 오름차순

      const nameA = a.name || '';
      const nameB = b.name || '';
      return nameA.localeCompare(nameB, 'ko');
    };

    const collectForSemester = (sem) => {
      const subjectMap = new Map(); // subject -> { subject, status, studentMap }

      gradeKeys.forEach(gKey => {
        this.ensureYearSemesterData(targetYear, gKey, sem);
        const semObj = (yearData[gKey] && yearData[gKey][sem]) || {};
        const offered = semObj.offeredSubjects || [];
        const enrollments = semObj.courseEnrollments || {};
        const masters = semObj.masterStudents || [];
        const masterMap = new Map(masters.map(s => [s.id, s]));
        const courseStatus = semObj.courseStatus || {};

        offered.forEach(sub => {
          if (!subjectMap.has(sub)) {
            subjectMap.set(sub, {
              subject: sub,
              status: courseStatus[sub] || 'confirmed',
              studentMap: new Map()
            });
          }
          const subInfo = subjectMap.get(sub);
          if (courseStatus[sub]) subInfo.status = courseStatus[sub];

          const studentIds = enrollments[sub] || [];
          studentIds.forEach(id => {
            const st = masterMap.get(id);
            if (st) {
              subInfo.studentMap.set(id, st);
            }
          });
        });
      });

      // 과목명 가나다순 정렬
      const sortedSubjects = Array.from(subjectMap.values()).sort((a, b) => a.subject.localeCompare(b.subject, 'ko'));
      return sortedSubjects.map(item => {
        const studentList = Array.from(item.studentMap.values()).sort(compareStudents);
        return {
          subject: item.subject,
          status: item.status,
          count: studentList.length,
          students: studentList
        };
      });
    };

    const sem1Subjects = collectForSemester('1');
    const sem2Subjects = collectForSemester('2');

    return {
      academicYear: targetYear,
      schoolName: '강원온라인학교',
      sem1Subjects,
      sem2Subjects
    };
  }

  // ==========================================
  // 학기·과목별 수강자 명단 엑셀 다운로드
  // ==========================================
  exportSubjectRosterExcel() {
    const data = this.collectSubjectRosterData();
    const totalCount = data.sem1Subjects.reduce((acc, s) => acc + s.count, 0) + 
                       data.sem2Subjects.reduce((acc, s) => acc + s.count, 0);

    if (data.sem1Subjects.length === 0 && data.sem2Subjects.length === 0) {
      this.showToast('등록된 개설 과목이 없습니다.');
      return;
    }

    ExcelService.exportSubjectRosterExcel(data);
    this.showToast(`[${data.academicYear}학년도] 학기·과목별 수강자 명단 엑셀이 다운로드되었습니다. (총 ${totalCount}명 수강 편제)`, true);
  }

  // ==========================================
  // 학기·과목별 수강자 명단 모달 보기 제어
  // ==========================================
  openSubjectRosterModal() {
    this.subjectRosterFilterSem = 'all'; // 'all', '1', '2'
    this.subjectRosterSearch = '';
    const modal = document.getElementById('modal-subject-roster');
    if (!modal) return;

    const searchInput = document.getElementById('subject-roster-search-input');
    if (searchInput) searchInput.value = '';

    modal.classList.remove('hidden');
    modal.classList.add('flex');
    this.renderSubjectRosterModal();
  }

  closeSubjectRosterModal() {
    const modal = document.getElementById('modal-subject-roster');
    if (!modal) return;
    modal.classList.add('hidden');
    modal.classList.remove('flex');
  }

  setSubjectRosterFilterSemester(sem) {
    this.subjectRosterFilterSem = sem;
    this.renderSubjectRosterModal();
  }

  handleSubjectRosterSearch(e) {
    this.subjectRosterSearch = (e.target.value || '').trim().toLowerCase();
    this.renderSubjectRosterModal();
  }

  renderSubjectRosterModal() {
    const data = this.collectSubjectRosterData();
    const container = document.getElementById('subject-roster-list-container');
    const badge = document.getElementById('subject-roster-stat-badge');
    const titleYear = document.getElementById('subject-roster-modal-year');
    if (!container) return;

    if (titleYear) {
      titleYear.textContent = `${data.academicYear}학년도`;
    }

    const totalSem1Count = data.sem1Subjects.reduce((acc, s) => acc + s.count, 0);
    const totalSem2Count = data.sem2Subjects.reduce((acc, s) => acc + s.count, 0);

    if (badge) {
      badge.innerHTML = `
        <span class="px-2 py-0.5 rounded-md bg-blue-100 text-blue-800 font-bold">1학기 ${data.sem1Subjects.length}개 과목 (${totalSem1Count}명)</span>
        <span class="px-2 py-0.5 rounded-md bg-purple-100 text-purple-800 font-bold">2학기 ${data.sem2Subjects.length}개 과목 (${totalSem2Count}명)</span>
      `;
    }

    // 탭 버튼 스타일 갱신
    ['all', '1', '2'].forEach(t => {
      const btn = document.getElementById(`tab-roster-sem-${t}`);
      if (btn) {
        if (this.subjectRosterFilterSem === t) {
          btn.className = 'px-3 py-1 text-xs font-bold rounded-lg bg-blue-600 text-white shadow-xs transition';
        } else {
          btn.className = 'px-3 py-1 text-xs font-semibold rounded-lg bg-white text-slate-600 hover:bg-slate-100 border border-slate-200 transition';
        }
      }
    });

    const q = (this.subjectRosterSearch || '').toLowerCase();

    // 과목 카드 렌더링 헬퍼
    const renderSubjectBlock = (semNum, subjects, theme) => {
      let filtered = subjects;
      if (q) {
        filtered = subjects.filter(s => {
          if (s.subject.toLowerCase().includes(q)) return true;
          return (s.students || []).some(st => {
            const name = (st.name || '').toLowerCase();
            const code = (st.studentCode || ExcelService.getStudentCode(st)).toLowerCase();
            return name.includes(q) || code.includes(q);
          });
        });
      }

      if (filtered.length === 0) {
        return `
          <div class="p-6 text-center text-xs text-slate-400 border border-dashed rounded-xl bg-slate-50/50">
            ${q ? `'${q}' 검색 조건에 맞는 과목이 없습니다.` : '등록된 과목이 없습니다.'}
          </div>
        `;
      }

      let subHtml = '<div class="space-y-4">';
      filtered.forEach((subItem, sIdx) => {
        const isClosed = subItem.status === 'canceled';
        const isAdjusted = subItem.status === 'adjusted';
        const statusBadge = isClosed 
          ? '<span class="text-[10px] px-2 py-0.5 rounded bg-red-100 text-red-700 font-bold">폐강</span>'
          : (isAdjusted ? '<span class="text-[10px] px-2 py-0.5 rounded bg-amber-100 text-amber-800 font-bold">조정</span>' : '<span class="text-[10px] px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold">정상</span>');

        subHtml += `
          <div class="border rounded-xl overflow-hidden shadow-2xs ${theme.cardBorder} bg-white">
            <!-- 과목 헤더 바 -->
            <div class="px-4 py-2.5 flex items-center justify-between gap-3 ${theme.headerBg} border-b ${theme.cardBorder}">
              <div class="flex items-center gap-2 min-w-0">
                <span class="w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold ${theme.seqBadge}">
                  ${sIdx + 1}
                </span>
                <span class="text-sm font-bold ${theme.titleColor} truncate">
                  ${subItem.subject}
                </span>
                ${statusBadge}
              </div>
              <div class="flex items-center gap-2 shrink-0">
                <span class="text-xs font-bold px-2.5 py-1 rounded-full ${theme.countBadge}">
                  수강 <strong class="text-sm">${subItem.count}</strong>명
                </span>
              </div>
            </div>

            <!-- 학생 목록 테이블 -->
            <div class="p-3">
              ${subItem.students.length === 0 ? `
                <div class="py-3 text-center text-xs text-slate-400">수강 신청 학생이 없습니다.</div>
              ` : `
                <div class="overflow-x-auto">
                  <table class="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr class="text-[11px] font-bold text-slate-500 border-b border-slate-200 bg-slate-50/70">
                        <th class="py-1.5 px-2.5 w-12 text-center">번호</th>
                        <th class="py-1.5 px-2.5 w-16 text-center">학년</th>
                        <th class="py-1.5 px-2.5 w-16 text-center">반</th>
                        <th class="py-1.5 px-2.5 w-16 text-center">번호</th>
                        <th class="py-1.5 px-2.5 w-24 text-center">학번</th>
                        <th class="py-1.5 px-3">성명</th>
                        <th class="py-1.5 px-2.5 text-right text-slate-400">비고</th>
                      </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100">
                      ${subItem.students.map((st, idx) => {
                        const code = st.studentCode || ExcelService.getStudentCode(st);
                        const matchHighlight = q && ((st.name || '').toLowerCase().includes(q) || code.toLowerCase().includes(q));
                        return `
                          <tr class="${matchHighlight ? 'bg-amber-50 font-bold' : (idx % 2 === 1 ? 'bg-slate-50/40' : '')} hover:bg-blue-50/50 transition">
                            <td class="py-1.5 px-2.5 text-center text-slate-400 text-[11px]">${idx + 1}</td>
                            <td class="py-1.5 px-2.5 text-center font-bold text-slate-800">${st.grade || '1'}학년</td>
                            <td class="py-1.5 px-2.5 text-center text-slate-700">${st.classNum || '1'}반</td>
                            <td class="py-1.5 px-2.5 text-center text-slate-700">${st.studentNum || '-'}번</td>
                            <td class="py-1.5 px-2.5 text-center font-mono font-semibold text-slate-600">${code}</td>
                            <td class="py-1.5 px-3 font-bold text-slate-900">${st.name || '-'}</td>
                            <td class="py-1.5 px-2.5 text-right text-[11px] text-emerald-600 font-medium">수강</td>
                          </tr>
                        `;
                      }).join('')}
                    </tbody>
                  </table>
                </div>
              `}
            </div>
          </div>
        `;
      });
      subHtml += '</div>';
      return subHtml;
    };

    let fullHtml = '';

    // 1학기 섹션
    if (this.subjectRosterFilterSem === 'all' || this.subjectRosterFilterSem === '1') {
      const sem1Theme = {
        cardBorder: 'border-blue-200',
        headerBg: 'bg-blue-50/80',
        seqBadge: 'bg-blue-600 text-white',
        titleColor: 'text-blue-950',
        countBadge: 'bg-blue-100 text-blue-900 border border-blue-200'
      };

      fullHtml += `
        <div class="mb-8">
          <div class="flex items-center gap-2 mb-3 pb-2 border-b-2 border-blue-400">
            <span class="px-2.5 py-1 rounded-lg bg-blue-600 text-white text-xs font-bold flex items-center gap-1 shadow-xs">
              <i data-lucide="calendar" class="w-3.5 h-3.5"></i>
              <span>1학기</span>
            </span>
            <h3 class="text-sm font-bold text-blue-950">1학기 과목별 수강자 명단</h3>
            <span class="text-xs text-slate-500 ml-auto font-medium">총 ${data.sem1Subjects.length}개 과목 · ${totalSem1Count}명</span>
          </div>
          ${renderSubjectBlock('1', data.sem1Subjects, sem1Theme)}
        </div>
      `;
    }

    // 2학기 섹션
    if (this.subjectRosterFilterSem === 'all' || this.subjectRosterFilterSem === '2') {
      const sem2Theme = {
        cardBorder: 'border-purple-200',
        headerBg: 'bg-purple-50/80',
        seqBadge: 'bg-purple-600 text-white',
        titleColor: 'text-purple-950',
        countBadge: 'bg-purple-100 text-purple-900 border border-purple-200'
      };

      fullHtml += `
        <div>
          <div class="flex items-center gap-2 mb-3 pb-2 border-b-2 border-purple-400">
            <span class="px-2.5 py-1 rounded-lg bg-purple-600 text-white text-xs font-bold flex items-center gap-1 shadow-xs">
              <i data-lucide="calendar" class="w-3.5 h-3.5"></i>
              <span>2학기</span>
            </span>
            <h3 class="text-sm font-bold text-purple-950">2학기 과목별 수강자 명단</h3>
            <span class="text-xs text-slate-500 ml-auto font-medium">총 ${data.sem2Subjects.length}개 과목 · ${totalSem2Count}명</span>
          </div>
          ${renderSubjectBlock('2', data.sem2Subjects, sem2Theme)}
        </div>
      `;
    }

    container.innerHTML = fullHtml;
    if (window.lucide) window.lucide.createIcons();
  }

  /**
   * 임의의 학기 데이터(semData)로부터 flat student 목록을 완벽하게 생성하는 헬퍼
   */
  buildFlatStudentListFromSemesterData(semData) {
    if (!semData) return [];
    const list = [];
    const masterStudents = semData.masterStudents || [];
    const courseEnrollments = semData.courseEnrollments || {};
    const courseSchedules = semData.courseSchedules || {};
    const offeredSubjects = semData.offeredSubjects || [];
    const periods = semData.periods || [];

    const masterMap = new Map(masterStudents.map(s => [s.id, s]));

    // 1. 과목별 편성 및 시간표 슬롯 매핑
    offeredSubjects.forEach(subject => {
      const studentIds = courseEnrollments[subject] || [];
      const schedules = courseSchedules[subject] || [];

      if (schedules.length === 0) {
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
        schedules.forEach(slot => {
          const periodObj = periods.find(p => p.id === slot.periodId || (!slot.periodId && p.period === slot.period));
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

    // 2. 미편제 학생도 등록
    const enrolledIds = new Set();
    Object.values(courseEnrollments).forEach(arr => arr.forEach(id => enrolledIds.add(id)));
    masterStudents.forEach(m => {
      if (!enrolledIds.has(m.id)) {
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

    return list;
  }

  // ==========================================
  // 실행 취소(되돌리기 / Undo) & 다시 실행(재실행 / Redo) 시스템
  // ==========================================
  pushUndoSnapshot(description = '작업 변경') {
    try {
      this.saveActiveYearSemesterData();
      const snapshot = {
        description: description,
        timestamp: Date.now(),
        activeYear: this.activeYear,
        activeSemester: this.activeSemester,
        academicYears: JSON.parse(JSON.stringify(this.academicYears)),
        yearsData: JSON.parse(JSON.stringify(this.yearsData))
      };
      this.undoStack.push(snapshot);
      if (this.undoStack.length > this.maxUndoDepth) {
        this.undoStack.shift();
      }
      // 새로운 작업이 수행되면 다시 실행(Redo) 스택은 초기화
      this.redoStack = [];
      this.updateUndoRedoButtonUI();
    } catch (e) {
      console.error('Failed to create undo snapshot:', e);
    }
  }

  handleUndo() {
    if (!this.undoStack || this.undoStack.length === 0) {
      this.showToast('되돌릴 이전 작업 내역이 없습니다.');
      return;
    }

    // 현재 상태를 다시 실행(Redo) 스택에 보관
    this.saveActiveYearSemesterData();
    const currentSnapshot = {
      description: this.undoStack[this.undoStack.length - 1]?.description || '되돌리기 전 상태',
      timestamp: Date.now(),
      activeYear: this.activeYear,
      activeSemester: this.activeSemester,
      academicYears: JSON.parse(JSON.stringify(this.academicYears)),
      yearsData: JSON.parse(JSON.stringify(this.yearsData))
    };
    if (!this.redoStack) this.redoStack = [];
    this.redoStack.push(currentSnapshot);
    if (this.redoStack.length > this.maxUndoDepth) {
      this.redoStack.shift();
    }

    const snapshot = this.undoStack.pop();
    if (!snapshot) return;

    this.activeYear = snapshot.activeYear;
    this.activeSemester = snapshot.activeSemester;
    this.academicYears = snapshot.academicYears;
    this.yearsData = snapshot.yearsData;

    this.loadActiveYearSemesterData();
    this.sortPeriods();
    this.buildSubjectColorMap();
    this.saveState();
    this.render();
    this.updateUndoRedoButtonUI();

    this.showToast(`[실행 취소 완료] '${snapshot.description}' 이전 상태로 되돌렸습니다. (재실행 가능: Ctrl+Y)`, false);
  }

  handleRedo() {
    if (!this.redoStack || this.redoStack.length === 0) {
      this.showToast('다시 실행할 내역이 없습니다.');
      return;
    }

    // 현재 상태를 되돌리기(Undo) 스택에 보관
    this.saveActiveYearSemesterData();
    const currentSnapshot = {
      description: '다시 실행 전 상태',
      timestamp: Date.now(),
      activeYear: this.activeYear,
      activeSemester: this.activeSemester,
      academicYears: JSON.parse(JSON.stringify(this.academicYears)),
      yearsData: JSON.parse(JSON.stringify(this.yearsData))
    };
    if (!this.undoStack) this.undoStack = [];
    this.undoStack.push(currentSnapshot);
    if (this.undoStack.length > this.maxUndoDepth) {
      this.undoStack.shift();
    }

    const snapshot = this.redoStack.pop();
    if (!snapshot) return;

    this.activeYear = snapshot.activeYear;
    this.activeSemester = snapshot.activeSemester;
    this.academicYears = snapshot.academicYears;
    this.yearsData = snapshot.yearsData;

    this.loadActiveYearSemesterData();
    this.sortPeriods();
    this.buildSubjectColorMap();
    this.saveState();
    this.render();
    this.updateUndoRedoButtonUI();

    this.showToast(`[다시 실행 완료] '${snapshot.description}' 작업을 다시 적용했습니다.`, false);
  }

  updateUndoRedoButtonUI() {
    // 1. 실행 취소(되돌리기) 버튼
    const undoBtn = document.getElementById('btn-app-undo');
    const undoBadge = document.getElementById('app-undo-count-badge');
    if (undoBtn) {
      const undoCount = this.undoStack ? this.undoStack.length : 0;
      if (undoCount > 0) {
        const lastAction = this.undoStack[this.undoStack.length - 1].description;
        undoBtn.disabled = false;
        undoBtn.className = 'px-2.5 py-1.5 text-xs font-bold rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300 shadow-xs transition flex items-center gap-1.5 cursor-pointer shrink-0';
        undoBtn.title = `클릭 시 실행 취소: '${lastAction}' 이전 상태로 복구 (Ctrl+Z)`;
        if (undoBadge) {
          undoBadge.textContent = `${undoCount}`;
          undoBadge.classList.remove('hidden');
        }
      } else {
        undoBtn.disabled = true;
        undoBtn.className = 'px-2.5 py-1.5 text-xs font-medium rounded-lg bg-slate-100 text-slate-400 border border-slate-200 transition flex items-center gap-1.5 cursor-not-allowed opacity-60 shadow-2xs shrink-0';
        undoBtn.title = '되돌릴 이전 작업 내역이 없습니다 (Ctrl+Z)';
        if (undoBadge) undoBadge.classList.add('hidden');
      }
    }

    // 2. 다시 실행(재실행) 버튼
    const redoBtn = document.getElementById('btn-app-redo');
    const redoBadge = document.getElementById('app-redo-count-badge');
    if (redoBtn) {
      const redoCount = this.redoStack ? this.redoStack.length : 0;
      if (redoCount > 0) {
        const nextAction = this.redoStack[this.redoStack.length - 1].description;
        redoBtn.disabled = false;
        redoBtn.className = 'px-2.5 py-1.5 text-xs font-bold rounded-lg bg-blue-50 hover:bg-blue-100 text-blue-900 border border-blue-300 shadow-xs transition flex items-center gap-1.5 cursor-pointer shrink-0';
        redoBtn.title = `클릭 시 다시 실행: '${nextAction}' 적용 (Ctrl+Y, Ctrl+Shift+Z)`;
        if (redoBadge) {
          redoBadge.textContent = `${redoCount}`;
          redoBadge.classList.remove('hidden');
        }
      } else {
        redoBtn.disabled = true;
        redoBtn.className = 'px-2.5 py-1.5 text-xs font-medium rounded-lg bg-slate-100 text-slate-400 border border-slate-200 transition flex items-center gap-1.5 cursor-not-allowed opacity-60 shadow-2xs shrink-0';
        redoBtn.title = '다시 실행할 내역이 없습니다 (Ctrl+Y, Ctrl+Shift+Z)';
        if (redoBadge) redoBadge.classList.add('hidden');
      }
    }
  }

  updateUndoButtonUI() {
    this.updateUndoRedoButtonUI();
  }

  // ==========================================
  // 렌더링 총괄
  // ==========================================
  render() {
    this.updateStats();
    this.renderAcademicYearTabs();
    this.renderAcademicGradeTabs();
    this.updateWeekendToggleBtn();
    this.updateUndoButtonUI();

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

  // 대상 학년 (2학년 / 3학년) 탭 렌더링
  renderAcademicGradeTabs() {
    const container = document.getElementById('academic-grade-tabs');
    if (!container) return;

    let html = '';
    ['2', '3'].forEach(g => {
      const isActive = this.activeGrade === g;
      const count = this.getStudentCountForGrade(g);
      html += `
        <button type="button" 
                onclick="window.app.switchGrade('${g}')"
                class="px-3.5 py-1.5 rounded-lg text-xs font-black flex items-center gap-1.5 transition ${
                  isActive 
                    ? 'bg-gradient-to-r from-indigo-600 to-purple-600 text-white shadow-xs ring-2 ring-indigo-400' 
                    : 'bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold'
                }"
                title="${this.activeYear}학년도 ${g}학년 작업 환경으로 전환">
          <span>${g}학년</span>
          <span class="text-[10px] px-1.5 py-0.2 rounded-full font-bold ${
            isActive ? 'bg-white text-indigo-700 shadow-2xs' : 'bg-slate-200 text-slate-600'
          }">
            ${count}명
          </span>
        </button>
      `;
    });
    container.innerHTML = html;
  }

  getStudentCountForGrade(grade) {
    const yearObj = this.yearsData[this.activeYear] || {};
    const gradeObj = yearObj[grade] || {};
    const semObj = gradeObj[this.activeSemester] || gradeObj['1'] || {};
    return (semObj.masterStudents && Array.isArray(semObj.masterStudents)) ? semObj.masterStudents.length : 0;
  }

  renderAcademicYearTabs() {
    const container = document.getElementById('academic-year-tabs');
    const badgeLabel = document.getElementById('current-year-label');
    const titleYear = document.getElementById('timetable-academic-year-text');
    const printYear = document.getElementById('timetable-print-year-text');

    const fullYearSemText = `${this.activeYear}학년도 ${this.activeGrade}학년 ${this.activeSemester}학기`;
    if (badgeLabel) badgeLabel.textContent = fullYearSemText;
    if (titleYear) titleYear.textContent = fullYearSemText;
    if (printYear) printYear.textContent = fullYearSemText;

    if (container) {
      container.innerHTML = this.academicYears.map(year => {
        const isActive = year === this.activeYear;
        const yearObj = this.yearsData[year] || {};
        
        let totalCount = 0;
        ['2', '3'].forEach(g => {
          const gObj = yearObj[g] || {};
          ['1', '2'].forEach(s => {
            const semObj = gObj[s] || {};
            if (semObj.masterStudents && Array.isArray(semObj.masterStudents)) {
              totalCount += semObj.masterStudents.length;
            }
          });
        });

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
    const currentGradeObj = currentYearObj[this.activeGrade] || {};
    const sem1Count = (currentGradeObj['1'] && currentGradeObj['1'].masterStudents) ? currentGradeObj['1'].masterStudents.length : 0;
    const sem2Count = (currentGradeObj['2'] && currentGradeObj['2'].masterStudents) ? currentGradeObj['2'].masterStudents.length : 0;

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

  // ==========================================
  // [학생 명부 엑셀 풀 & 성명 검색 & 선택 등록]
  // ==========================================
  loadExcelStudentPool() {
    try {
      const saved = localStorage.getItem(this.excelPoolStorageKey);
      if (saved) {
        this.excelStudentPool = JSON.parse(saved);
      }
    } catch (e) {
      console.warn('Failed to load excel student pool:', e);
      this.excelStudentPool = [];
    }
    // 풀이 아직 비어있고 현재 등록된 학생(masterStudents)이 있다면 초기 풀로 자동 등록
    if ((!this.excelStudentPool || this.excelStudentPool.length === 0) && this.masterStudents && this.masterStudents.length > 0) {
      this.excelStudentPool = this.masterStudents.map(s => ({
        grade: s.grade,
        classNum: s.classNum,
        studentNum: s.studentNum,
        studentCode: s.studentCode || ExcelService.getStudentCode(s),
        name: s.name
      }));
      this.saveExcelStudentPool();
    }
    this.updateExcelPoolUI();
  }

  saveExcelStudentPool() {
    try {
      localStorage.setItem(this.excelPoolStorageKey, JSON.stringify(this.excelStudentPool || []));
    } catch (e) {
      console.error('Failed to save excel student pool:', e);
    }
    this.updateExcelPoolUI();
  }

  updateExcelPoolUI() {
    const statusEl = document.getElementById('excel-pool-status');
    if (!statusEl) return;
    const count = (this.excelStudentPool || []).length;
    if (count > 0) {
      statusEl.textContent = `기준 데이터 유지 중 (${count}명 등록됨)`;
      statusEl.className = 'text-[10px] text-emerald-700 font-bold';
    } else {
      statusEl.textContent = '기준 명단 파일을 업로드해주세요';
      statusEl.className = 'text-[10px] text-indigo-700 font-medium';
    }
  }

  // ==========================================
  // [1-A. 기준 학생명부 (무결한 마스터 풀) 전용 업로드]
  // ==========================================
  async handleMasterRosterUpload(file) {
    if (!file) return;

    try {
      const parsed = await ExcelService.parseStudentExcel(file);
      if (!parsed || parsed.length === 0) {
        alert('학생명부 파일에서 학생 데이터를 찾을 수 없습니다.');
        return;
      }

      // 1. 실행 취소 스냅샷 저장
      this.pushUndoSnapshot('기준 학생명부 등록');

      // 2. 무결한 마스터 풀로 등록 (성명 & 학번 기준 중복 제거)
      const rosterMap = new Map();
      parsed.forEach(st => {
        const code = st.studentCode || ExcelService.getStudentCode(st);
        const key = `${code}_${st.name}`;
        rosterMap.set(key, {
          grade: st.grade,
          classNum: st.classNum,
          studentNum: st.studentNum,
          studentCode: code,
          name: st.name
        });
      });

      this.excelStudentPool = Array.from(rosterMap.values());
      this.saveExcelStudentPool();

      // 3. 현재 학기의 masterStudents와도 동기화
      if (!this.masterStudents || this.masterStudents.length === 0) {
        this.masterStudents = this.excelStudentPool.map(s => ({
          id: `sm_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
          grade: s.grade,
          classNum: s.classNum,
          studentNum: s.studentNum,
          studentCode: s.studentCode,
          name: s.name
        }));
      } else {
        // 이미 등록된 학생들의 학적을 무결한 명부 정보로 최신화
        this.masterStudents.forEach(ms => {
          const matched = this.excelStudentPool.find(p => p.name === ms.name);
          if (matched) {
            ms.grade = matched.grade;
            ms.classNum = matched.classNum;
            ms.studentNum = matched.studentNum;
            ms.studentCode = matched.studentCode;
          }
        });
      }

      this.rebuildStudentAssignments();
      this.saveState();
      this.render();

      this.showToast(`🛡️ 무결한 기준 학생명부(${this.excelStudentPool.length}명) 등록 완료! 앞으로 업로드되는 모든 시간표 엑셀의 기준 데이터로 영구 유지됩니다.`, true);

      // 검색창에 포커스
      const searchInput = document.getElementById('input-search-student-pool');
      if (searchInput) searchInput.focus();
    } catch (err) {
      alert('학생명부 등록 실패: ' + err.message);
    }
  }

  handleClearExcelStudentPool() {
    if (!this.excelStudentPool || this.excelStudentPool.length === 0) {
      this.showToast('등록된 기준 학생명부가 없습니다.');
      return;
    }
    if (confirm(`현재 등록된 기준 학생명부(${this.excelStudentPool.length}명)를 초기화하시겠습니까?\n\n※ 초기화 후에도 새로운 학생명부를 다시 업로드하실 수 있습니다.`)) {
      this.excelStudentPool = [];
      this.saveExcelStudentPool();
      this.showToast('기준 학생명부가 초기화되었습니다.');
    }
  }

  // ==========================================
  // [1-B. 시간표 통합 엑셀 업로드 (학기 구분 분기 반영)]
  // ==========================================
  async handleUnifiedExcelUpload(filesInput) {
    if (!filesInput) return;
    const fileList = Array.isArray(filesInput) 
      ? filesInput 
      : (filesInput instanceof FileList ? Array.from(filesInput) : [filesInput]);
    if (fileList.length === 0) return;

    try {
      // 1. 실행 취소 스냅샷 저장
      this.pushUndoSnapshot('통합 엑셀 업로드 반영');

      // 2. 무결한 학생명부 풀 (this.excelStudentPool) 인덱스 구성 (성명 기준)
      const poolByName = new Map();
      (this.excelStudentPool || []).forEach(p => {
        const clean = String(p.name || '').trim().replace(/\s+/g, '');
        if (!clean) return;
        if (!poolByName.has(clean)) poolByName.set(clean, []);
        poolByName.get(clean).push(p);
      });

      const hasMasterPool = (this.excelStudentPool || []).length > 0;

      let totalParsed = 0;
      let successFiles = 0;
      let exactCount = 0;
      const reconciledList = []; // 명부 기준으로 학년/반/번호가 자동 보정된 학생들
      const unmatchedList = [];  // 명부에 없어서 엑셀 기재 정보로 신규 등록된 학생들

      // 학기별 데이터 버킷: key = `${targetYear}_${targetSemester}` -> { year, semester, students: [], schedules: [] }
      const termBuckets = new Map();

      for (const file of fileList) {
        try {
          const parsed = await ExcelService.parseStudentExcel(file);
          if (parsed && parsed.length > 0) {
            
            parsed.forEach(st => {
              const rawName = String(st.name || '').trim();
              const cleanName = rawName.replace(/\s+/g, '');
              const candidates = poolByName.get(cleanName) || [];
              let matchedPool = null;

              if (candidates.length === 1) {
                matchedPool = candidates[0];
              } else if (candidates.length > 1) {
                matchedPool = candidates.find(c => {
                  if (st.studentCode && c.studentCode && st.studentCode === c.studentCode) return true;
                  if (st.grade === c.grade && st.classNum === c.classNum && st.studentNum === c.studentNum) return true;
                  if (st.grade === c.grade && st.classNum === c.classNum) return true;
                  return false;
                }) || candidates[0];
              }

              // 신청 과목 추출
              const subs = (st.subjects && Array.isArray(st.subjects) && st.subjects.length > 0)
                ? st.subjects
                : (st.subject ? [st.subject] : []);
              const validSubs = subs.map(s => String(s).trim()).filter(s => s && s !== '-' && s !== '미배정' && s !== '미입력');

              // 학생명부 대조 & 학년/반/번호 자동 보정
              if (matchedPool) {
                const origGrade = st.grade;
                const origClass = st.classNum;
                const origNum = st.studentNum;
                const origCode = st.studentCode;

                const isMismatch = (origGrade && origGrade !== matchedPool.grade) ||
                                   (origClass && origClass !== matchedPool.classNum) ||
                                   (origNum && matchedPool.studentNum && origNum !== matchedPool.studentNum);

                const isMissingInfo = !origGrade || !origClass || !origNum;

                if (isMismatch || isMissingInfo) {
                  reconciledList.push({
                    name: rawName,
                    origInfo: `${origGrade || '?'}학년 ${origClass || '?'}반 ${origNum || '-'}번`,
                    correctedInfo: `${matchedPool.grade}학년 ${matchedPool.classNum}반 ${matchedPool.studentNum || '-'}번`,
                    origCode: origCode || '-',
                    correctedCode: matchedPool.studentCode || ExcelService.getStudentCode(matchedPool)
                  });
                } else {
                  exactCount++;
                }

                // 무결한 학생명부의 정보로 자동 보정 및 최신화
                st.grade = matchedPool.grade;
                st.classNum = matchedPool.classNum;
                st.studentNum = matchedPool.studentNum;
                st.studentCode = matchedPool.studentCode || ExcelService.getStudentCode(matchedPool);
              } else {
                // 학생명부에 없는 학생 -> "해당 학생은 등록은 되기는 하는 거야."
                if (hasMasterPool) {
                  unmatchedList.push({
                    name: rawName,
                    info: `${st.grade || 1}학년 ${st.classNum || 1}반 ${st.studentNum || '-'}번`,
                    code: st.studentCode || ExcelService.getStudentCode(st),
                    subjects: validSubs.length > 0 ? validSubs.join(', ') : '과목 미배정'
                  });
                }
              }

              // 학년도 및 학기 구분 추출 (엑셀에 기재된 학기 구분 적용)
              const tYear = st.year || this.activeYear;
              const tSemester = (st.semester === '2' || String(st.semester).includes('2')) 
                ? '2' 
                : ((st.semester === '1' || String(st.semester).includes('1')) ? '1' : this.activeSemester);

              const termKey = `${tYear}_${tSemester}`;
              if (!termBuckets.has(termKey)) {
                termBuckets.set(termKey, {
                  year: tYear,
                  semester: tSemester,
                  records: []
                });
              }
              termBuckets.get(termKey).records.push({
                st,
                validSubs,
                schedules: st.schedules || []
              });

              totalParsed++;
            });
            successFiles++;
          }
        } catch (fileErr) {
          console.warn(`파일(${file.name}) 처리 실패:`, fileErr);
        }
      }

      // 3. 학년도 및 학기별 분기 저장 실행 (해당 학년도/학기의 기존 데이터는 비우고 새 파일 내용으로 완전 갱신)
      const termSummaryList = [];

      for (const [termKey, bucket] of termBuckets.entries()) {
        const { year: bYear, semester: bSemester, records } = bucket;

        // 해당 학년도/학기 데이터 구조 보장
        if (!this.academicYears.includes(bYear)) {
          this.academicYears.push(bYear);
          this.academicYears.sort((a, b) => Number(a) - Number(b));
        }
        this.ensureYearSemesterData(bYear, bSemester);
        const semData = this.yearsData[bYear][bSemester];

        // ★ 요청 사항 반영: 해당 학년도·학기의 기존 데이터 초기화 (새로운 파일의 내용으로 전면 갱신)
        semData.masterStudents = [];
        semData.courseEnrollments = {};
        semData.courseSchedules = {};
        semData.courseStatus = {};
        semData.offeredSubjects = [];
        semData.students = [];

        const semMasterMap = new Map();

        let semEnrolled = 0;
        let semScheduled = 0;

        records.forEach(({ st, validSubs, schedules }) => {
          const code = st.studentCode || ExcelService.getStudentCode(st);
          const key = `${code}_${st.name}`;

          // masterStudents 등록
          let mStudent = semMasterMap.get(key);
          if (!mStudent) {
            mStudent = {
              id: `sm_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
              grade: st.grade,
              classNum: st.classNum,
              studentNum: st.studentNum,
              studentCode: code,
              name: st.name
            };
            semMasterMap.set(key, mStudent);
            if (!semData.masterStudents) semData.masterStudents = [];
            semData.masterStudents.push(mStudent);
          } else {
            if (!mStudent.studentNum && st.studentNum) mStudent.studentNum = st.studentNum;
            if (!mStudent.grade && st.grade) mStudent.grade = st.grade;
            if (!mStudent.classNum && st.classNum) mStudent.classNum = st.classNum;
          }

          // 과목 바구니 편제
          if (!semData.offeredSubjects) semData.offeredSubjects = [];
          if (!semData.courseEnrollments) semData.courseEnrollments = {};
          if (!semData.courseStatus) semData.courseStatus = {};

          validSubs.forEach(sub => {
            if (!semData.offeredSubjects.includes(sub)) {
              semData.offeredSubjects.push(sub);
              semData.courseStatus[sub] = 'open';
            }
            if (!semData.courseEnrollments[sub]) {
              semData.courseEnrollments[sub] = [];
            }
            if (!semData.courseEnrollments[sub].includes(mStudent.id)) {
              semData.courseEnrollments[sub].push(mStudent.id);
              semEnrolled++;
            }
          });

          // 시간표 슬롯 배치
          if (!semData.courseSchedules) semData.courseSchedules = {};
          schedules.forEach(sch => {
            const sub = sch.subject;
            if (!sub || sub === '-' || sub === '미배정' || sub === '미입력') return;

            if (sch.courseStatus) {
              semData.courseStatus[sub] = sch.courseStatus;
            }

            if (sch.day && (sch.period || sch.periodName || sch.startTime)) {
              let periodObj = (semData.periods || []).find(p => {
                if (sch.period && !p.isBreak && p.period === sch.period) return true;
                if (sch.periodName && p.name === sch.periodName) return true;
                if (sch.startTime && p.startTime === sch.startTime) return true;
                return false;
              });

              if (!periodObj && (sch.period || sch.periodName)) {
                const newPid = `p_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`;
                periodObj = {
                  id: newPid,
                  period: sch.period || ((semData.periods || []).filter(p => !p.isBreak).length + 1),
                  name: sch.periodName || `${sch.period}교시`,
                  startTime: sch.startTime || '09:00',
                  endTime: sch.endTime || '09:50',
                  isBreak: false
                };
                if (!semData.periods) semData.periods = [];
                semData.periods.push(periodObj);
              }

              if (sch.day === 'sat' || sch.day === 'sun') {
                semData.showWeekends = true;
              }

              if (periodObj) {
                if (!semData.courseSchedules[sub]) semData.courseSchedules[sub] = [];
                const exists = semData.courseSchedules[sub].some(s => 
                  s.day === sch.day && (s.periodId === periodObj.id || (!s.periodId && s.period === periodObj.period))
                );
                if (!exists) {
                  semData.courseSchedules[sub].push({
                    id: `sch_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
                    day: sch.day,
                    periodId: periodObj.id,
                    period: periodObj.period,
                    status: sch.scheduleStatus || 'confirmed'
                  });
                  semScheduled++;
                }
              }
            }
          });
        });

        termSummaryList.push(`${bYear}년 ${bSemester}학기: 학생 ${semMasterMap.size}명 (${semEnrolled}건 편제)`);
      }

      // 4. 기준 학생명부 풀(검색용)도 미등록 학생 동기화
      const poolMap = new Map((this.excelStudentPool || []).map(s => {
        const c = s.studentCode || ExcelService.getStudentCode(s);
        return [`${c}_${s.name}`, s];
      }));
      (this.masterStudents || []).forEach(ms => {
        const c = ms.studentCode || ExcelService.getStudentCode(ms);
        if (!poolMap.has(`${c}_${ms.name}`)) {
          poolMap.set(`${c}_${ms.name}`, {
            grade: ms.grade,
            classNum: ms.classNum,
            studentNum: ms.studentNum,
            studentCode: c,
            name: ms.name
          });
        }
      });
      this.excelStudentPool = Array.from(poolMap.values());
      this.saveExcelStudentPool();

      // 5. 현재 활성 학년도/학기 데이터 리로드 & 렌더링
      this.loadActiveYearSemesterData();
      this.sortPeriods();
      this.buildSubjectColorMap();
      this.saveState();
      this.switchWorkflowTab('step1');
      this.render();

      // 6. 대조 결과 및 학기 구분 반영 알림 표시
      if (reconciledList.length > 0 || unmatchedList.length > 0) {
        this.openStudentReconcileModal({
          total: totalParsed,
          exact: exactCount,
          reconciledList: reconciledList,
          unmatchedList: unmatchedList
        });
      } else {
        const summaryStr = termSummaryList.length > 0 ? termSummaryList.join(', ') : '시간표 데이터';
        this.showToast(`🎉 통합 엑셀 갱신 완료: [${summaryStr}] 기존 데이터를 새 파일 내용으로 완전 갱신하였습니다.`, true);
      }

      // 검색창에 포커스
      const searchInput = document.getElementById('input-search-student-pool');
      if (searchInput) searchInput.focus();
    } catch (err) {
      alert('통합 엑셀 처리 실패: ' + err.message);
    }
  }

  // 기존 호환성 유지 메서드
  async handleExcelPoolUpload(filesInput) {
    return this.handleUnifiedExcelUpload(filesInput);
  }

  openStudentReconcileModal({ total, exact, reconciledList, unmatchedList }) {
    const modal = document.getElementById('student-reconcile-modal');
    if (!modal) return;

    // 통계 표시
    const elTotal = document.getElementById('reconcile-stat-total');
    const elExact = document.getElementById('reconcile-stat-exact');
    const elReconciled = document.getElementById('reconcile-stat-reconciled');
    const elUnmatched = document.getElementById('reconcile-stat-unmatched');

    if (elTotal) elTotal.textContent = `${total || 0}명`;
    if (elExact) elExact.textContent = `${exact || 0}명`;
    if (elReconciled) elReconciled.textContent = `${reconciledList?.length || 0}명`;
    if (elUnmatched) elUnmatched.textContent = `${unmatchedList?.length || 0}명`;

    // 1. 보정 목록 바인딩
    const secCorrected = document.getElementById('reconcile-section-corrected');
    const badgeCorrected = document.getElementById('reconcile-badge-corrected');
    const tbodyCorrected = document.getElementById('reconcile-table-corrected');

    if (reconciledList && reconciledList.length > 0) {
      if (secCorrected) secCorrected.classList.remove('hidden');
      if (badgeCorrected) badgeCorrected.textContent = `${reconciledList.length}명`;
      if (tbodyCorrected) {
        tbodyCorrected.innerHTML = reconciledList.map(item => `
          <tr class="hover:bg-amber-50/40 transition">
            <td class="py-2 px-3 font-bold text-slate-800">${item.name}</td>
            <td class="py-2 px-3 text-slate-500 line-through decoration-red-400">
              <span class="px-1.5 py-0.5 rounded bg-slate-100 text-slate-600 text-[10px]">${item.origInfo}</span>
            </td>
            <td class="py-2 px-3 text-center text-amber-500 font-bold">➔</td>
            <td class="py-2 px-3 font-bold text-emerald-700">
              <span class="px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 text-[11px] inline-flex items-center gap-1 font-bold">
                <i data-lucide="check" class="w-3 h-3 text-emerald-600"></i> ${item.correctedInfo}
              </span>
            </td>
          </tr>
        `).join('');
      }
    } else {
      if (secCorrected) secCorrected.classList.add('hidden');
    }

    // 2. 미등록 목록 바인딩
    const secUnmatched = document.getElementById('reconcile-section-unmatched');
    const badgeUnmatched = document.getElementById('reconcile-badge-unmatched');
    const tbodyUnmatched = document.getElementById('reconcile-table-unmatched');

    if (unmatchedList && unmatchedList.length > 0) {
      if (secUnmatched) secUnmatched.classList.remove('hidden');
      if (badgeUnmatched) badgeUnmatched.textContent = `${unmatchedList.length}명`;
      if (tbodyUnmatched) {
        tbodyUnmatched.innerHTML = unmatchedList.map(item => `
          <tr class="hover:bg-blue-50/40 transition">
            <td class="py-2 px-3 font-bold text-slate-800">${item.name}</td>
            <td class="py-2 px-3 text-slate-700 font-medium">${item.info}</td>
            <td class="py-2 px-3 text-slate-600 truncate max-w-xs" title="${item.subjects}">${item.subjects}</td>
            <td class="py-2 px-3 text-center">
              <span class="px-2 py-0.5 rounded-full bg-blue-100 text-blue-800 text-[10px] font-bold">
                정상 등록 완료 (신규)
              </span>
            </td>
          </tr>
        `).join('');
      }
    } else {
      if (secUnmatched) secUnmatched.classList.add('hidden');
    }

    modal.classList.remove('hidden');
    modal.classList.add('flex');
    if (window.lucide) window.lucide.createIcons();
  }

  closeStudentReconcileModal() {
    const modal = document.getElementById('student-reconcile-modal');
    if (modal) {
      modal.classList.add('hidden');
      modal.classList.remove('flex');
    }
  }

  handleSearchPoolInput(rawQuery) {
    const q = (rawQuery || '').trim();
    const dropdown = document.getElementById('pool-search-results-dropdown');
    const clearBtn = document.getElementById('btn-clear-pool-search');

    if (clearBtn) {
      if (q) clearBtn.classList.remove('hidden');
      else clearBtn.classList.add('hidden');
    }

    if (!dropdown) return;
    if (!q) {
      dropdown.classList.add('hidden');
      dropdown.innerHTML = '';
      return;
    }

    const lowerQ = q.toLowerCase();
    const matches = (this.excelStudentPool || []).filter(s => {
      const name = (s.name || '').toLowerCase();
      const code = (s.studentCode || '').toLowerCase();
      return name.includes(lowerQ) || code.includes(lowerQ);
    });

    this.renderPoolSearchResults(matches, q);
  }

  handleSearchPoolEnter(rawQuery) {
    const q = (rawQuery || '').trim().toLowerCase();
    if (!q) return;
    const matches = (this.excelStudentPool || []).filter(s => {
      const name = (s.name || '').toLowerCase();
      const code = (s.studentCode || '').toLowerCase();
      return name.includes(q) || code.includes(q);
    });

    // 아직 등록되지 않은 첫 번째 학생 바로 등록
    const unimported = matches.find(s => !this.isMasterStudentEnrolled(s));
    if (unimported) {
      this.importStudentFromPool(unimported);
    }
  }

  isMasterStudentEnrolled(poolStudent) {
    const pCode = poolStudent.studentCode || ExcelService.getStudentCode(poolStudent);
    return this.masterStudents.some(ms => {
      const mCode = ms.studentCode || ExcelService.getStudentCode(ms);
      return mCode === pCode && ms.name === poolStudent.name;
    });
  }

  renderPoolSearchResults(matches, query) {
    const dropdown = document.getElementById('pool-search-results-dropdown');
    if (!dropdown) return;

    if (matches.length === 0) {
      dropdown.innerHTML = `
        <div class="p-3 text-center text-xs text-slate-500">
          <p class="font-medium">'${query}' 검색 결과가 없습니다.</p>
          <p class="text-[10px] text-slate-400 mt-0.5">엑셀 명부에 없으면 상단의 <strong>[수기 등록]</strong>을 이용하세요.</p>
        </div>
      `;
      dropdown.classList.remove('hidden');
      return;
    }

    // 최대 30명 표출
    const slice = matches.slice(0, 30);
    const html = slice.map(s => {
      const isEnrolled = this.isMasterStudentEnrolled(s);
      const code = s.studentCode || ExcelService.getStudentCode(s);
      const encoded = encodeURIComponent(JSON.stringify(s));

      return `
        <div class="px-3 py-2 flex items-center justify-between gap-2 hover:bg-blue-50/70 transition cursor-pointer"
             onclick="window.app.handlePoolItemClick('${encoded}')">
          <div class="min-w-0 flex items-center gap-2">
            <span class="font-mono text-xs font-bold text-blue-700 bg-blue-50 px-1.5 py-0.5 rounded border border-blue-200 shrink-0">
              ${code}
            </span>
            <div class="min-w-0">
              <span class="font-bold text-xs text-slate-900">${s.name}</span>
              <span class="text-[11px] text-slate-500 ml-1.5">(${s.grade}학년 ${s.classNum}반 ${s.studentNum || '-'}번)</span>
            </div>
          </div>

          <div class="shrink-0">
            ${isEnrolled ? `
              <span class="inline-flex items-center gap-1 text-[11px] font-bold px-2 py-0.5 rounded bg-slate-100 text-slate-500">
                <i data-lucide="check" class="w-3 h-3 text-emerald-600"></i> 등록완료
              </span>
            ` : `
              <button type="button" 
                      onclick="event.stopPropagation(); window.app.handlePoolItemClick('${encoded}')" 
                      class="px-2.5 py-1 text-xs font-bold bg-blue-600 hover:bg-blue-700 text-white rounded-md shadow-2xs transition flex items-center gap-1">
                <i data-lucide="plus" class="w-3 h-3"></i>
                <span>선택 등록</span>
              </button>
            `}
          </div>
        </div>
      `;
    }).join('');

    dropdown.innerHTML = html;
    dropdown.classList.remove('hidden');
    if (window.lucide) window.lucide.createIcons();
  }

  handlePoolItemClick(encodedStudent) {
    try {
      const student = JSON.parse(decodeURIComponent(encodedStudent));
      this.importStudentFromPool(student);
    } catch (e) {
      console.error(e);
    }
  }

  clearPoolSearch() {
    const inputSearch = document.getElementById('input-search-student-pool');
    const dropdown = document.getElementById('pool-search-results-dropdown');
    const clearBtn = document.getElementById('btn-clear-pool-search');

    if (inputSearch) {
      inputSearch.value = '';
    }
    if (clearBtn) {
      clearBtn.classList.add('hidden');
    }
    if (dropdown) {
      dropdown.classList.add('hidden');
      dropdown.innerHTML = '';
    }
  }

  importStudentFromPool(poolStudent) {
    if (this.isMasterStudentEnrolled(poolStudent)) {
      this.showToast(`'${poolStudent.name}' 학생은 이미 등록되어 있습니다.`);
      return;
    }

    const code = poolStudent.studentCode || ExcelService.getStudentCode(poolStudent);
    const newStudent = {
      id: `sm_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
      grade: poolStudent.grade,
      classNum: poolStudent.classNum,
      studentNum: poolStudent.studentNum,
      studentCode: code,
      name: poolStudent.name
    };

    this.masterStudents.push(newStudent);
    this.rebuildStudentAssignments();
    this.saveState();
    this.render();

    // 학생 등록 완료 후 검색창에 남아있던 검색어 및 추천 드롭다운 즉시 초기화(삭제)
    this.clearPoolSearch();

    this.showToast(`'${code} ${poolStudent.name}' 학생이 시간표 편제 목록에 등록되었습니다.`);

    // 다음 학생을 편리하게 연속 검색할 수 있도록 검색창 포커스
    const inputSearch = document.getElementById('input-search-student-pool');
    if (inputSearch) {
      inputSearch.focus();
    }
  }

  // ==========================================
  // [학생 명단 정렬 (학번순 / 이름순 / 과목명순)]
  // ==========================================
  setStudentSort(sortKey) {
    this.studentSortKey = sortKey;

    const keys = ['code', 'name', 'subject'];
    keys.forEach(k => {
      const btn = document.getElementById(`sort-btn-${k}`);
      if (btn) {
        if (k === sortKey) {
          btn.className = 'px-2 py-0.5 rounded text-[11px] font-bold bg-blue-600 text-white shadow-2xs transition';
        } else {
          btn.className = 'px-2 py-0.5 rounded text-[11px] font-medium bg-white text-slate-700 border border-slate-300 hover:bg-slate-50 transition';
        }
      }
    });

    this.renderStep1StudentList();
  }

  getFilteredMasterStudents() {
    const filtered = this.masterStudents.filter(student => {
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

    // 정렬: 1순위 - 미편성(미배정) 학생 최상단 배치 (사용자 요청 반영)
    filtered.sort((a, b) => {
      const subsA = this.getStudentEnrolledCourses(a.id);
      const subsB = this.getStudentEnrolledCourses(b.id);
      const isUnassignedA = subsA.length === 0 ? 0 : 1;
      const isUnassignedB = subsB.length === 0 ? 0 : 1;

      // 미편성 학생(0)이 편성 완료된 학생(1)보다 항상 상단에 먼저 오도록 정렬
      if (isUnassignedA !== isUnassignedB) {
        return isUnassignedA - isUnassignedB;
      }

      // 2순위: 선택된 정렬 기준 적용 (학번순 / 이름순 / 과목명순)
      const codeA = a.studentCode || ExcelService.getStudentCode(a);
      const codeB = b.studentCode || ExcelService.getStudentCode(b);

      if (this.studentSortKey === 'name') {
        const nameCompare = (a.name || '').localeCompare(b.name || '', 'ko');
        if (nameCompare !== 0) return nameCompare;
        return codeA.localeCompare(codeB, undefined, { numeric: true });
      }

      if (this.studentSortKey === 'subject') {
        const subAStr = subsA.length > 0 ? subsA.sort().join(', ') : '';
        const subBStr = subsB.length > 0 ? subsB.sort().join(', ') : '';
        const subCompare = subAStr.localeCompare(subBStr, 'ko');
        if (subCompare !== 0) return subCompare;
        return (a.name || '').localeCompare(b.name || '', 'ko');
      }

      // 기본: 학번순 (code)
      return codeA.localeCompare(codeB, undefined, { numeric: true });
    });

    return filtered;
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

            <!-- 학번 & 성명 (클릭하여 학생 정보 수정) -->
            <div class="flex items-baseline gap-1.5 min-w-0 cursor-pointer group"
                 onclick="event.stopPropagation(); window.app.openEditStudentModal('${st.id}')"
                 title="클릭하여 학생 정보(학번·이름·수강과목) 수정">
              <span class="font-mono font-bold text-xs text-slate-900">${code}</span>
              <span class="font-bold text-xs text-slate-800 truncate group-hover:text-blue-600">${st.name}</span>
              <i data-lucide="edit-2" class="w-2.5 h-2.5 text-slate-300 group-hover:text-blue-500"></i>
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
              <span class="text-[10px] px-2 py-0.5 rounded-full bg-amber-50 text-amber-700 border border-amber-200 font-bold inline-flex items-center gap-1">
                <span class="w-1.5 h-1.5 rounded-full bg-amber-500"></span> 미편성
              </span>
            `}

            <!-- 학생 정보 수정 버튼 -->
            <button type="button" 
                    onclick="event.stopPropagation(); window.app.openEditStudentModal('${st.id}')" 
                    class="p-1 text-slate-400 hover:text-blue-600 rounded transition ml-1" 
                    title="학생 정보 및 수강 과목 수정">
              <i data-lucide="edit-3" class="w-3.5 h-3.5"></i>
            </button>

            <!-- 학생 삭제 버튼 -->
            <button type="button" 
                    onclick="event.stopPropagation(); window.app.deleteMasterStudent('${st.id}')" 
                    class="p-1 text-slate-300 hover:text-red-500 rounded transition" 
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
    this.sortOfferedSubjects();
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

    if (addedCount > 0) {
      this.pushUndoSnapshot(`학생 ${addedCount}명 '${targetSubject}' 바구니 담기`);
    }

    this.selectedStep1StudentIds.clear();
    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`선택한 ${addedCount}명의 학생이 '${targetSubject}' 바구니에 담겼습니다.`, addedCount > 0);
  }

  deleteMasterStudent(studentId, skipConfirm = false) {
    const st = this.masterStudents.find(s => s.id === studentId);
    if (!st) return;
    if (!skipConfirm && !confirm(`'${st.studentCode} ${st.name}' 학생을 명단에서 삭제하시겠습니까?\n(모든 과목 편제 및 시간표에서도 제외됩니다)`)) {
      return;
    }

    this.pushUndoSnapshot(`학생 '${st.name}' 삭제`);

    this.masterStudents = this.masterStudents.filter(s => s.id !== studentId);
    this.selectedStep1StudentIds.delete(studentId);

    // 모든 과목 편제에서도 제외
    Object.keys(this.courseEnrollments).forEach(sub => {
      this.courseEnrollments[sub] = this.courseEnrollments[sub].filter(id => id !== studentId);
    });

    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`${st.name} 학생이 삭제되었습니다.`, true);
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

    this.pushUndoSnapshot(`과목 '${subjectName}' 바구니 삭제`);

    this.offeredSubjects = this.offeredSubjects.filter(s => s !== subjectName);
    delete this.courseEnrollments[subjectName];
    delete this.courseSchedules[subjectName];

    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`'${subjectName}' 과목 바구니가 삭제되었습니다.`, true);
  }

  // ==========================================
  // 과목 상태 관리 (개설 / 조정 / 폐강)
  // ==========================================
  getSubjectStatus(subjectName) {
    if (!this.courseStatus) this.courseStatus = {};
    return this.courseStatus[subjectName] || 'open';
  }

  cycleSubjectStatus(subjectName) {
    if (!this.courseStatus) this.courseStatus = {};
    const current = this.getSubjectStatus(subjectName);
    let next = 'open';
    if (current === 'open') next = 'adjust';
    else if (current === 'adjust') next = 'closed';
    else next = 'open';

    this.courseStatus[subjectName] = next;
    this.saveState();
    this.render();
    const labels = { open: '개설', adjust: '조정', closed: '폐강' };
    this.showToast(`'${subjectName}' 과목이 [${labels[next]}] 상태로 변경되었습니다.`);
  }

  setSubjectStatus(subjectName, status) {
    if (!this.courseStatus) this.courseStatus = {};
    this.courseStatus[subjectName] = status;
    this.saveState();
    this.render();
    const labels = { open: '개설', adjust: '조정', closed: '폐강' };
    this.showToast(`'${subjectName}' 과목이 [${labels[status] || status}] 상태로 설정되었습니다.`);
  }

  renderSubjectBuckets() {
    this.sortOfferedSubjects();
    const container = document.getElementById('subject-buckets-container');
    const sideListContainer = document.getElementById('side-subject-list-container');
    const sideCountBadge = document.getElementById('side-subject-count-badge');

    if (sideCountBadge) {
      sideCountBadge.textContent = `${this.offeredSubjects.length}개`;
    }

    // [좌측]: 세로 과목 목록 렌더링 (과목명, 인원수, 개설/조정/폐강 상태 배지 표시 및 클릭 스크롤)
    if (sideListContainer) {
      if (this.offeredSubjects.length === 0) {
        sideListContainer.innerHTML = `
          <div class="py-6 text-center text-slate-400 text-xs italic">
            개설된 과목이 없습니다.
          </div>
        `;
      } else {
        sideListContainer.innerHTML = this.offeredSubjects.map(sub => {
          const color = this.getSubjectColor(sub);
          const studentIds = this.courseEnrollments[sub] || [];
          const safeSub = encodeURIComponent(sub);
          const count = studentIds.length;
          const status = this.getSubjectStatus(sub);
          const isSelected = this.activeSelectedSubject === sub;

          // 상태별 배지 스타일 및 텍스트
          let statusBadgeClass = 'bg-emerald-50 text-emerald-700 border-emerald-300 hover:bg-emerald-100 hover:border-emerald-400';
          let statusSymbol = '🟢';
          let statusLabel = '개설';
          let itemStyleClass = 'border-slate-200 bg-white hover:border-indigo-400 hover:bg-indigo-50/40';

          if (status === 'adjust') {
            statusBadgeClass = 'bg-amber-50 text-amber-700 border-amber-300 hover:bg-amber-100 hover:border-amber-400';
            statusSymbol = '▲';
            statusLabel = '조정';
            itemStyleClass = 'border-amber-200 bg-amber-50/30 hover:border-amber-400 hover:bg-amber-50/60';
          } else if (status === 'closed') {
            statusBadgeClass = 'bg-rose-50 text-rose-700 border-rose-300 hover:bg-rose-100 hover:border-rose-400';
            statusSymbol = '✖️';
            statusLabel = '폐강';
            itemStyleClass = 'border-slate-300 bg-slate-100/90 text-slate-400';
          }

          if (isSelected) {
            itemStyleClass += ' ring-2 ring-indigo-300 border-indigo-500';
          }

          return `
            <div id="side-item-${safeSub}"
                 class="side-subject-item group flex items-center justify-between gap-1.5 px-2.5 py-2 rounded-xl border transition-all cursor-pointer shadow-2xs ${itemStyleClass}"
                 draggable="false"
                 onclick="window.app.scrollToBucket(decodeURIComponent('${safeSub}'))"
                 ondragover="window.app.handleBucketDragOver(event)"
                 ondragleave="window.app.handleBucketDragLeave(event)"
                 ondrop="window.app.handleBucketDrop(event, decodeURIComponent('${safeSub}'))"
                 title="클릭하여 해당 과목 바구니로 이동 / 학생을 여기에 드롭하여 편제">
              <div class="flex items-center gap-1.5 min-w-0">
                <span class="w-2.5 h-2.5 rounded-full shrink-0 shadow-2xs" 
                      style="background-color: ${status === 'closed' ? '#94A3B8' : (status === 'adjust' ? '#F59E0B' : color.text)};"></span>
                <span class="text-xs font-bold whitespace-nowrap truncate ${status === 'closed' ? 'line-through text-slate-400' : 'text-slate-800 group-hover:text-indigo-900'}">
                  ${sub}
                </span>
              </div>

              <div class="flex items-center gap-1 shrink-0">
                <!-- 과목명 수정 버튼 -->
                <button type="button" 
                        onclick="event.stopPropagation(); window.app.handleRenameSubject(decodeURIComponent('${safeSub}'))" 
                        class="p-1 rounded text-slate-300 hover:text-indigo-600 hover:bg-slate-100 transition" 
                        title="'${sub}' 과목명 수정">
                  <i data-lucide="edit-3" class="w-3 h-3"></i>
                </button>

                <!-- 인원수 배지 -->
                <span class="text-[10px] font-black px-1.5 py-0.5 rounded-full shadow-2xs ${
                  status === 'closed' ? 'bg-slate-200 text-slate-500' : (status === 'adjust' ? 'bg-amber-100 text-amber-800' : (count > 0 ? 'bg-indigo-100 text-indigo-700' : 'bg-slate-100 text-slate-400'))
                }">
                  (${count}명)
                </span>

                <!-- 개설 / 조정 / 폐강 3단 순환 토글 배지 버튼 -->
                <button type="button" 
                        onclick="event.stopPropagation(); window.app.cycleSubjectStatus(decodeURIComponent('${safeSub}'))"
                        class="px-1.5 py-0.5 text-[10px] font-black rounded border transition-all shadow-2xs flex items-center gap-0.5 ${statusBadgeClass}"
                        title="클릭하여 상태 순환 (개설 ➔ 조정 ➔ 폐강 ➔ 개설)">
                  <span class="text-[9px]">${statusSymbol}</span>
                  <span>${statusLabel}</span>
                </button>
              </div>
            </div>
          `;
        }).join('');
      }
    }

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
      const safeSub = encodeURIComponent(sub);
      const status = this.getSubjectStatus(sub);

      let cardBorder = 'border-slate-200/90';
      let headerBg = color.bg;
      let headerBorder = color.border;

      if (status === 'adjust') {
        cardBorder = 'border-amber-300 border-dashed';
        headerBg = '#FEF3C7';
        headerBorder = '#FDE68A';
      } else if (status === 'closed') {
        cardBorder = 'border-slate-300 opacity-90';
        headerBg = '#F1F5F9';
        headerBorder = '#CBD5E1';
      }

      html += `
        <div id="bucket-card-${safeSub}" 
             class="bucket-card w-full shrink-0 bg-white rounded-2xl border-2 shadow-sm flex flex-col overflow-hidden transition-all hover:shadow-lg scroll-mt-36 ${cardBorder}"
             ondragover="window.app.handleBucketDragOver(event)"
             ondragleave="window.app.handleBucketDragLeave(event)"
             ondrop="window.app.handleBucketDrop(event, decodeURIComponent('${safeSub}'))">
          
          <!-- 바구니 상단 헤더 -->
          <div class="p-4 border-b flex items-center justify-between shrink-0"
               style="background-color: ${headerBg}; border-color: ${headerBorder};">
            <div class="flex items-center gap-2.5 flex-wrap">
              <span class="w-4 h-4 rounded-full border shadow-2xs" 
                    style="background-color: ${status === 'closed' ? '#94A3B8' : (status === 'adjust' ? '#F59E0B' : color.text)}; border-color: ${headerBorder};"></span>
              <div class="flex items-center gap-1.5">
                <h3 class="font-black text-base ${status === 'closed' ? 'line-through text-slate-500' : ''}" style="color: ${status === 'closed' ? '#64748B' : color.text};">${sub}</h3>
                <button type="button" 
                        onclick="event.stopPropagation(); window.app.handleRenameSubject(decodeURIComponent('${safeSub}'))" 
                        class="p-1 rounded-md text-slate-400 hover:text-indigo-600 hover:bg-white/90 transition shadow-2xs" 
                        title="'${sub}' 과목명 수정">
                  <i data-lucide="edit-3" class="w-3.5 h-3.5"></i>
                </button>
              </div>

              <!-- 과목명 오른쪽: 학생 이름 검색 및 엔터 즉시 등록 입력창 -->
              <div class="flex items-center bg-white/95 border border-indigo-200 rounded-lg px-2 py-0.5 shadow-2xs focus-within:ring-2 focus-within:ring-indigo-400 focus-within:border-indigo-400 transition" 
                   onclick="event.stopPropagation()">
                <i data-lucide="user-plus" class="w-3.5 h-3.5 text-indigo-500 mr-1 shrink-0"></i>
                <input type="text" 
                       id="input-enroll-student-${safeSub}" 
                       placeholder="학생명 입력 후 Enter" 
                       class="text-xs bg-transparent border-none outline-none w-28 sm:w-36 text-slate-800 placeholder:text-slate-400 font-medium py-0.5" 
                       onkeydown="if(event.key === 'Enter') { event.preventDefault(); window.app.handleQuickEnrollStudentByEnter(decodeURIComponent('${safeSub}'), this); }" 
                       title="학생 이름을 입력하고 Enter를 누르면 기준 학생명부에서 찾아 '${sub}' 과목에 즉시 등록됩니다.">
                <button type="button" 
                        onclick="event.stopPropagation(); const inp = document.getElementById('input-enroll-student-${safeSub}'); if(inp) window.app.handleQuickEnrollStudentByEnter(decodeURIComponent('${safeSub}'), inp);" 
                        class="ml-1 text-indigo-600 hover:text-indigo-800 p-0.5 rounded transition cursor-pointer" 
                        title="등록">
                  <i data-lucide="corner-down-left" class="w-3 h-3"></i>
                </button>
              </div>

              <span class="text-xs font-bold px-2.5 py-0.5 rounded-full bg-white/90 border text-slate-800 shadow-2xs">
                ${studentIds.length}명 수강
              </span>

              <!-- 개설 / 조정 / 폐강 원클릭 순환 배지 버튼 (선택지 없이 클릭할 때마다 순환 변경) -->
              <button type="button" 
                      onclick="event.stopPropagation(); window.app.cycleSubjectStatus(decodeURIComponent('${safeSub}'))" 
                      class="px-2.5 py-1 text-xs font-black rounded-lg border transition-all shadow-2xs flex items-center gap-1.5 ${
                        status === 'open' 
                          ? 'bg-emerald-100 text-emerald-800 border-emerald-300 hover:bg-emerald-200' 
                          : (status === 'adjust' 
                              ? 'bg-amber-100 text-amber-800 border-amber-300 hover:bg-amber-200' 
                              : 'bg-rose-100 text-rose-800 border-rose-300 hover:bg-rose-200')
                      }" 
                      title="클릭하여 상태 순환 (🟢 개설 ➔ 🔺 조정 ➔ ✖️ 폐강 ➔ 🟢 개설)">
                <span>${status === 'open' ? '🟢' : (status === 'adjust' ? '🔺' : '✖️')}</span>
                <span>${status === 'open' ? '개설' : (status === 'adjust' ? '조정' : '폐강')}</span>
              </button>
            </div>

            <div class="flex items-center gap-1.5">
              <!-- 학년도/학기 전용 이전/복사 버튼 -->
              <button type="button" 
                      onclick="window.app.openSubjectTransferModal(decodeURIComponent('${safeSub}'))" 
                      class="px-2.5 py-1 text-xs font-bold rounded-lg bg-white/90 hover:bg-white text-indigo-700 border border-indigo-200 shadow-2xs transition flex items-center gap-1" 
                      title="'${sub}' 과목을 다른 학기/학년도로 복사 또는 이동">
                <i data-lucide="arrow-left-right" class="w-3.5 h-3.5 text-indigo-600"></i>
                <span class="hidden sm:inline">학기 이전/복사</span>
              </button>

              <!-- 과목 바구니 삭제 -->
              <button type="button" 
                      onclick="window.app.handleDeleteSubjectBucket(decodeURIComponent('${safeSub}'))" 
                      class="text-slate-400 hover:text-red-600 p-1.5 rounded-lg hover:bg-white/80 transition" title="과목 바구니 삭제">
                <i data-lucide="trash-2" class="w-4 h-4"></i>
              </button>
            </div>
          </div>

          <!-- 바구니 본문: 수강생 명단 & 시간표 설정 (드래그 영역 박스 제거) -->
          <div class="p-4 sm:p-5 flex flex-col gap-4">
            
            <!-- 편제된 학생 명단 칩 목록 -->
            <div>
              <div class="flex items-center justify-between text-xs font-bold text-slate-800 mb-2.5">
                <span class="flex items-center gap-1.5">
                  <i data-lucide="users" class="w-4 h-4 text-indigo-600"></i>
                  <span>수강 학생 명단 (${studentIds.length}명)</span>
                </span>
                ${studentIds.length > 0 ? `
                  <button type="button" onclick="window.app.handleClearBucketStudents(decodeURIComponent('${safeSub}'))" class="text-xs text-slate-400 hover:text-red-600 font-medium transition flex items-center gap-1">
                    <i data-lucide="rotate-ccw" class="w-3 h-3"></i>
                    <span>전체 비우기</span>
                  </button>
                ` : ''}
              </div>

              <div class="flex flex-wrap content-start gap-2 min-h-[64px] max-h-72 overflow-y-auto p-3.5 bg-slate-50/90 rounded-xl border border-slate-200">
                ${studentIds.length > 0 ? studentIds.map(stId => {
                  const m = masterMap.get(stId);
                  if (!m) return '';
                  const code = m.studentCode || ExcelService.getStudentCode(m);
                  return `
                    <span class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white border border-slate-200 text-slate-800 shadow-2xs hover:border-blue-400 hover:shadow-xs transition cursor-pointer group" 
                          onclick="window.app.openEditStudentModal('${m.id}')" 
                          title="클릭하여 학생 정보(학번·이름·수강과목) 수정">
                      <span class="font-mono text-[11px] text-slate-500 font-bold">${code}</span>
                      <span class="font-bold text-slate-900 group-hover:text-blue-600">${m.name}</span>
                      <i data-lucide="edit-2" class="w-2.5 h-2.5 text-slate-300 group-hover:text-blue-500"></i>
                      <button type="button" 
                              onclick="event.stopPropagation(); window.app.handleRemoveStudentFromCourse(decodeURIComponent('${safeSub}'), '${m.id}')" 
                              class="text-slate-400 hover:text-red-600 ml-1 text-sm font-bold leading-none p-0.5" title="수강 제외">&times;</button>
                    </span>
                  `;
                }).join('') : `
                  <div class="text-xs text-slate-400 py-3 px-4 flex items-center justify-center w-full gap-2">
                    <i data-lucide="arrow-left" class="w-4 h-4 text-indigo-400"></i>
                    <span>좌측 <strong>과목 목록</strong>으로 학생 카드를 드래그하여 편제하세요.</span>
                  </div>
                `}
              </div>
            </div>

            <!-- 과목 시간표 시간 설정 바 (시간표 자동 배치) -->
            <div class="pt-3.5 border-t border-slate-200 flex flex-col gap-2">
              <div class="flex items-center justify-between text-xs font-bold text-slate-800">
                <span class="flex items-center gap-1.5">
                  <i data-lucide="clock" class="w-4 h-4 text-indigo-600"></i>
                  <span>과목 시간 설정 (종합 시간표 자동 배치)</span>
                </span>
                <span class="text-xs text-indigo-700 font-bold px-2 py-0.5 rounded-full bg-indigo-50 border border-indigo-200">${schedules.length}시간 편성</span>
              </div>

              <!-- 요일 + 교시 추가 폼 -->
              <div class="flex items-center gap-2">
                <select id="bucket-day-${sub}" class="flex-1 px-3 py-2 text-xs bg-slate-50 border border-slate-300 rounded-lg font-medium text-slate-800 focus:ring-1 focus:ring-indigo-500">
                  ${activeDays.map(d => `<option value="${d.id}">${d.fullName}</option>`).join('')}
                </select>
                <select id="bucket-period-${sub}" class="flex-1 px-3 py-2 text-xs bg-slate-50 border border-slate-300 rounded-lg font-medium text-slate-800 focus:ring-1 focus:ring-indigo-500">
                  ${regularPeriods.map(p => `<option value="${p.id}">${p.name} (${p.startTime})</option>`).join('')}
                </select>
                <button type="button" 
                        onclick="window.app.handleAddBucketSchedule(decodeURIComponent('${safeSub}'))" 
                        class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-lg transition whitespace-nowrap shadow-xs flex items-center gap-1">
                  <i data-lucide="plus" class="w-3.5 h-3.5"></i>
                  <span>시간 추가</span>
                </button>
              </div>

              <!-- 현재 편성된 시간 칩 목록 (여러 날, 여러 시간) -->
              <div class="flex flex-wrap gap-1.5 mt-1 min-h-[30px] items-center">
                ${schedules.length > 0 ? schedules.map(sch => {
                  const dayName = ExcelService.dayIdToName(sch.day);
                  const periodObj = this.periods.find(p => p.id === sch.periodId || (!sch.periodId && p.period === sch.period));
                  const periodName = periodObj ? periodObj.name : `${sch.period}교시`;
                  const sym = ExcelService.getStatusSymbol(sch.status || 'confirmed');
                  return `
                    <span class="inline-flex items-center gap-1.5 px-3 py-1 rounded-lg text-xs font-bold bg-indigo-50 text-indigo-900 border border-indigo-200 shadow-2xs">
                      <span>${sym}</span>
                      <span>${dayName} ${periodName}</span>
                      <button type="button" 
                              onclick="window.app.handleRemoveBucketSchedule(decodeURIComponent('${safeSub}'), '${sch.id || (sch.day + '_' + sch.periodId)}')"
                              class="text-indigo-400 hover:text-red-600 ml-1 text-sm font-bold" title="시간표에서 배치 해제">&times;</button>
                    </span>
                  `;
                }).join('') : `
                  <span class="text-xs text-slate-400 italic">아직 시간표에 배치된 시간이 없습니다. 요일과 교시를 선택 후 추가하세요.</span>
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

  scrollToBucket(subjectName) {
    this.activeSelectedSubject = subjectName;

    // 1. 좌측 목록 아이템 활성 하이라이트 갱신
    document.querySelectorAll('.side-subject-item').forEach(el => {
      el.classList.remove('ring-2', 'ring-indigo-400', 'border-indigo-500', 'bg-indigo-50/70');
    });
    const sideItem = document.getElementById(`side-item-${encodeURIComponent(subjectName)}`);
    if (sideItem) {
      sideItem.classList.add('ring-2', 'ring-indigo-400', 'border-indigo-500', 'bg-indigo-50/70');
    }

    // 2. 우측 스크롤 컨테이너 내에서 해당 카드가 가장 위에 정확히 보이도록 정밀 스크롤
    const container = document.getElementById('subject-buckets-container');
    const safeSub = encodeURIComponent(subjectName);
    const card = document.getElementById(`bucket-card-${safeSub}`);

    if (container && card) {
      const containerRect = container.getBoundingClientRect();
      const cardRect = card.getBoundingClientRect();
      // 컨테이너 맨 위에서 딱 2px 여백을 두고 정확하게 최상단에 안착
      const offsetDiff = cardRect.top - containerRect.top;
      const targetScrollTop = Math.max(0, container.scrollTop + offsetDiff - 2);

      container.scrollTo({
        top: targetScrollTop,
        behavior: 'smooth'
      });

      // 3. 눈에 띄는 선명한 하이라이트 링 효과
      card.classList.remove('ring-4', 'ring-indigo-500', 'ring-offset-2');
      void card.offsetWidth; // reflow 트리거
      card.classList.add('ring-4', 'ring-indigo-500', 'ring-offset-2');
      setTimeout(() => {
        card.classList.remove('ring-4', 'ring-indigo-500', 'ring-offset-2');
      }, 1600);
    }
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

  // ==========================================
  // 과목 바구니: 기준 학생명부 기반 학생 검색 및 엔터 즉시 등록
  // ==========================================
  handleQuickEnrollStudentByEnter(subjectName, inputEl) {
    if (!inputEl) return;
    const rawVal = inputEl.value.trim();
    if (!rawVal) {
      this.showToast('등록할 학생의 이름을 입력하세요.');
      return;
    }

    const cleanQuery = rawVal.replace(/\s+/g, '').toLowerCase();

    // 1. 기준 학생명부(excelStudentPool)에서 정확 일치 검색
    let matches = (this.excelStudentPool || []).filter(s => {
      const sName = String(s.name || '').trim().replace(/\s+/g, '').toLowerCase();
      return sName === cleanQuery;
    });

    // 1-1. 정확 일치가 없다면 이름 포함 검색
    if (matches.length === 0) {
      matches = (this.excelStudentPool || []).filter(s => {
        const sName = String(s.name || '').trim().replace(/\s+/g, '').toLowerCase();
        return sName.includes(cleanQuery);
      });
    }

    // 1-2. 기준 학생명부에 없다면 현재 학기 등록 학생(masterStudents)에서 보조 검색
    if (matches.length === 0 && this.masterStudents && this.masterStudents.length > 0) {
      matches = this.masterStudents.filter(s => {
        const sName = String(s.name || '').trim().replace(/\s+/g, '').toLowerCase();
        return sName === cleanQuery || sName.includes(cleanQuery);
      });
    }

    // 일치하는 학생이 전혀 없는 경우
    if (matches.length === 0) {
      this.showToast(`기준 학생명부에서 '${rawVal}' 학생을 찾을 수 없습니다. 학생명부를 확인해 주세요.`);
      return;
    }

    // 동명이인 (2명 이상)일 경우 prompt로 선택
    let target = null;
    if (matches.length === 1) {
      target = matches[0];
    } else {
      const listStr = matches.map((m, idx) => {
        const code = m.studentCode || (window.ExcelService ? window.ExcelService.getStudentCode(m) : `${m.grade}${String(m.classNum).padStart(2, '0')}${String(m.studentNum || 0).padStart(2, '0')}`);
        return `${idx + 1}. ${m.name} (${m.grade}학년 ${m.classNum}반 ${m.studentNum || '-'}번 / 학번: ${code})`;
      }).join('\n');

      const ans = prompt(`'${rawVal}' 이름으로 여러 명의 학생이 검색되었습니다.\n등록할 학생의 번호를 입력하세요:\n\n${listStr}`, '1');
      if (!ans) return;
      const selectedIdx = parseInt(ans.trim(), 10) - 1;
      if (isNaN(selectedIdx) || selectedIdx < 0 || selectedIdx >= matches.length) {
        this.showToast('선택한 번호가 올바르지 않아 등록이 취소되었습니다.');
        return;
      }
      target = matches[selectedIdx];
    }

    // 2. 현재 활성 학기 masterStudents에 등록되어 있는지 확인 및 없으면 추가
    let masterStudent = null;
    const targetCode = target.studentCode || (window.ExcelService ? window.ExcelService.getStudentCode(target) : `${target.grade}${String(target.classNum).padStart(2, '0')}${String(target.studentNum || 0).padStart(2, '0')}`);

    if (target.id) {
      masterStudent = this.masterStudents.find(s => s.id === target.id);
    }
    if (!masterStudent) {
      masterStudent = this.masterStudents.find(s => {
        const mCode = s.studentCode || (window.ExcelService ? window.ExcelService.getStudentCode(s) : `${s.grade}${String(s.classNum).padStart(2, '0')}${String(s.studentNum || 0).padStart(2, '0')}`);
        return mCode === targetCode && s.name === target.name;
      });
    }

    if (!masterStudent) {
      masterStudent = {
        id: `sm_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`,
        grade: target.grade,
        classNum: target.classNum,
        studentNum: target.studentNum,
        studentCode: targetCode,
        name: target.name
      };
      this.masterStudents.push(masterStudent);
    }

    // 3. 해당 과목에 학생 등록
    if (!this.courseEnrollments[subjectName]) {
      this.courseEnrollments[subjectName] = [];
    }

    if (this.courseEnrollments[subjectName].includes(masterStudent.id)) {
      this.showToast(`'${masterStudent.name}' 학생은 이미 '${subjectName}' 과목에 등록되어 있습니다.`);
      inputEl.value = '';
      return;
    }

    this.pushUndoSnapshot(`'${subjectName}' 과목에 '${masterStudent.name}' 학생 등록`);
    this.courseEnrollments[subjectName].push(masterStudent.id);
    this.rebuildStudentAssignments();
    this.saveState();
    this.render();

    this.showToast(`'${masterStudent.name}' 학생이 '${subjectName}' 과목에 등록되었습니다!`, true);

    // 4. 화면 리렌더링 후 해당 과목 인풋에 다시 포커스를 주어 연속 입력 지원
    setTimeout(() => {
      const safeSub = encodeURIComponent(subjectName);
      const nextInput = document.getElementById(`input-enroll-student-${safeSub}`);
      if (nextInput) {
        nextInput.focus();
        nextInput.value = '';
      }
    }, 50);
  }

  handleRemoveStudentFromCourse(subjectName, studentId) {
    if (!this.courseEnrollments[subjectName]) return;
    this.pushUndoSnapshot(`'${subjectName}' 과목에서 학생 제외`);
    this.courseEnrollments[subjectName] = this.courseEnrollments[subjectName].filter(id => id !== studentId);
    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`학생이 '${subjectName}' 과목에서 제외되었습니다.`, true);
  }

  handleClearBucketStudents(subjectName) {
    if (!confirm(`'${subjectName}' 과목의 모든 수강생 편제를 비우시겠습니까?`)) return;
    this.pushUndoSnapshot(`'${subjectName}' 수강생 전체 비우기`);
    this.courseEnrollments[subjectName] = [];
    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`'${subjectName}' 과목의 편제 학생이 모두 비워졌습니다.`, true);
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
      const safeSub = encodeURIComponent(sub);
      const status = this.getSubjectStatus(sub);
      const isSelected = this.highlightedSubject === sub;
      const isDimmed = Boolean(this.highlightedSubject && !isSelected);

      html += `
        <div class="draggable-palette-subject p-3 rounded-xl border transition-all hover:scale-[1.01] shadow-2xs cursor-grab active:cursor-grabbing select-none ${
          status === 'closed' ? 'opacity-60 bg-slate-100 border-slate-300' : ''
        } ${
          isSelected 
            ? 'ring-2 ring-indigo-600 ring-offset-2 border-indigo-600 shadow-md scale-[1.02] bg-indigo-50/30' 
            : (isDimmed ? 'opacity-40 hover:opacity-80' : '')
        }"
             style="background-color: ${isSelected ? undefined : (status === 'closed' ? '#F8FAFC' : color.bg)}; border-color: ${isSelected ? '#4F46E5' : (status === 'closed' ? '#CBD5E1' : color.border)};"
             draggable="true"
             onclick="if(!window.app.isDraggingPalette) window.app.toggleHighlightSubject(decodeURIComponent('${safeSub}'))"
             ondragstart="window.app.handlePaletteDragStart(event, decodeURIComponent('${safeSub}'))"
             ondragend="window.app.handlePaletteDragEnd(event)"
             title="클릭: 오른쪽 시간표에서 이 과목 하이라이트 (토글) / 드래그: 시간표에 배치">
          
          <div class="flex items-center justify-between gap-1 mb-1.5 pointer-events-none">
            <div class="flex items-center gap-1.5 min-w-0">
              <i data-lucide="grip-vertical" class="w-3.5 h-3.5 shrink-0" style="color: ${status === 'closed' ? '#94A3B8' : color.text};"></i>
              <span class="font-bold text-xs truncate ${status === 'closed' ? 'line-through text-slate-400' : ''}" style="color: ${status === 'closed' ? '#64748B' : color.text};">${sub}</span>
            </div>
            <div class="flex items-center gap-1 pointer-events-auto">
              <!-- 원클릭 하이라이트 버튼 -->
              <button type="button" 
                      onclick="event.stopPropagation(); window.app.toggleHighlightSubject(decodeURIComponent('${safeSub}'))" 
                      class="px-2 py-0.5 text-[11px] font-black rounded-lg transition shadow-2xs cursor-pointer flex items-center gap-1 shrink-0 ${
                        isSelected 
                          ? 'bg-indigo-600 text-white hover:bg-indigo-700 ring-2 ring-indigo-400 ring-offset-1' 
                          : 'bg-white/95 text-indigo-700 hover:bg-indigo-50 border border-indigo-200'
                      }"
                      title="오른쪽 시간표에서 '${sub}' 과목 하이라이트 토글">
                <i data-lucide="sparkles" class="w-3 h-3 ${isSelected ? 'text-amber-300' : 'text-indigo-600'}"></i>
                <span>${isSelected ? '강조중' : '강조'}</span>
              </button>
              ${status === 'closed' ? `
                <span class="text-[10px] font-black px-1.5 py-0.2 rounded bg-rose-100 text-rose-700 border border-rose-300">✖️ 폐강</span>
              ` : (status === 'adjust' ? `
                <span class="text-[10px] font-black px-1.5 py-0.2 rounded bg-amber-100 text-amber-800 border border-amber-300">▲ 조정</span>
              ` : '')}
              <span class="text-[11px] font-bold px-2 py-0.5 rounded-full bg-white/90 border text-slate-700 shadow-2xs">
                ${studentIds.length}명
              </span>
            </div>
          </div>

          <div class="flex items-center justify-between text-[11px] text-slate-600 pointer-events-none">
            <span>
              ${schedules.length > 0 ? `
                <strong class="text-indigo-700 font-bold">${schedules.length}시간</strong> 배치됨
              ` : `
                <span class="text-amber-600 font-semibold">미배치</span>
              `}
            </span>
            <span class="text-[10px] text-slate-500 font-medium flex items-center gap-0.5">
              <span>시간표로 드래그</span> &rarr;
            </span>
          </div>

        </div>
      `;
    });

    container.innerHTML = html;
    if (window.lucide) window.lucide.createIcons();
  }

  // 과목 팔레트 선택 시 오른쪽 시간표 하이라이트 토글
  toggleHighlightSubject(subjectName) {
    if (this.highlightedSubject === subjectName) {
      this.highlightedSubject = null;
      this.showToast('과목 강조 표시가 해제되었습니다.');
    } else {
      this.highlightedSubject = subjectName;
      const count = (this.courseSchedules[subjectName] || []).length;
      this.showToast(`'${subjectName}' 과목이 시간표에서 강조 표시되었습니다. (총 ${count}시간 배치됨)`, true);
      
      // 첫 번째 강조 카드로 부드럽게 스크롤
      setTimeout(() => {
        const firstCard = document.querySelector('.subject-card.is-highlighted');
        if (firstCard) {
          firstCard.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'nearest' });
        }
      }, 80);
    }
    this.renderStep2Palette();
    this.renderTimetable();
  }

  // 과목 하이라이트 즉시 해제
  clearSubjectHighlight() {
    this.highlightedSubject = null;
    this.showToast('과목 강조 표시가 해제되었습니다.');
    this.renderStep2Palette();
    this.renderTimetable();
  }

  handlePaletteDragStart(e, subjectName) {
    this.isDraggingPalette = true;
    this.dragData = {
      type: 'palette-subject',
      subjectName: subjectName
    };
    try {
      e.dataTransfer.setData('text/plain', JSON.stringify(this.dragData));
      e.dataTransfer.setData('application/json', JSON.stringify(this.dragData));
    } catch (err) {}
    e.dataTransfer.effectAllowed = 'copyMove';
    e.currentTarget.classList.add('is-dragging');
  }

  handlePaletteDragEnd(e) {
    this.dragData = null;
    setTimeout(() => {
      this.isDraggingPalette = false;
    }, 120);
    if (e.currentTarget) e.currentTarget.classList.remove('is-dragging');
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

          // 현재 팔레트에서 선택/하이라이트된 과목이 이 셀에 배치되어 있는지 여부
          const isCellHighlighted = Boolean(
            this.highlightedSubject && 
            scheduledEntries.some(e => e.subject === this.highlightedSubject)
          );

          bodyHtml += `
            <td class="timetable-cell p-1.5 border-l border-slate-200 align-top transition-colors ${
              day.isWeekend ? 'bg-red-50/20' : ''
            } ${isCellHighlighted ? 'has-highlight' : ''}"
                data-day="${day.id}"
                data-period-id="${period.id}"
                ondragover="window.app.handleTimetableCellDragOver(event)"
                ondragleave="window.app.handleTimetableCellDragLeave(event)"
                ondrop="window.app.handleTimetableCellDrop(event, '${day.id}', '${period.id}')">
              
              <div class="min-h-[72px] flex flex-col gap-1.5">
          `;

          if (scheduledEntries.length > 0) {
            scheduledEntries.forEach(({ subject, schedule }) => {
              const color = this.getSubjectColor(subject);
              const enrolledIds = this.courseEnrollments[subject] || [];
              const status = schedule.status || 'confirmed';
              const symHtml = ExcelService.getStatusHtml(status);
              const isCanceled = status === 'canceled';
              const isAdjust = status === 'adjust';
              const safeSub = subject.replace(/\\/g, '\\\\').replace(/'/g, "\\'");

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

              // 동일 시간(day.id, period.id)에 다른 유효 과목에도 편성된 학생 검사 (시간 중복)
              const conflictStudents = [];
              if (!isCanceled && enrolledIds.length > 0) {
                enrolledIds.forEach(stId => {
                  const overlappingSubjects = [];
                  scheduledEntries.forEach(other => {
                    if (other.subject === subject) return;
                    const otherStatus = other.schedule.status || 'confirmed';
                    if (otherStatus === 'canceled') return; // 폐강 과목 제외
                    const otherEnrolledIds = this.courseEnrollments[other.subject] || [];
                    if (otherEnrolledIds.includes(stId)) {
                      overlappingSubjects.push(other.subject);
                    }
                  });

                  if (overlappingSubjects.length > 0) {
                    const st = masterMap.get(stId);
                    const stName = st ? st.name : '학생';
                    const stCode = st ? (st.studentCode || ExcelService.getStudentCode(st)) : '';
                    conflictStudents.push({
                      name: stName,
                      code: stCode,
                      subjects: overlappingSubjects
                    });
                  }
                });
              }

              const conflictCount = conflictStudents.length;
              let conflictTooltip = '';
              if (conflictCount > 0) {
                const lines = [
                  `⚠️ 시간 중복 수강생 (${conflictCount}명):`,
                  ...conflictStudents.map(c => `• ${c.name}${c.code ? `(${c.code})` : ''}: [${c.subjects.join(', ')}] 과목과 동시 배정`)
                ];
                conflictTooltip = lines.join('\n').replace(/"/g, '&quot;');
              }

              const isCardHighlighted = Boolean(this.highlightedSubject && this.highlightedSubject === subject);
              const isCardDimmed = Boolean(this.highlightedSubject && this.highlightedSubject !== subject);

              bodyHtml += `
                <div class="subject-card p-2 rounded-lg border text-xs transition-all duration-200 relative select-none cursor-grab active:cursor-grabbing ${
                  isCanceled ? 'is-canceled' : ''
                } ${isAdjust ? 'is-adjust' : ''} ${
                  isCardHighlighted 
                    ? 'is-highlighted font-black' 
                    : (isCardDimmed ? 'is-dimmed' : '')
                }"
                     style="background-color: ${isCanceled ? '#E2E8F0' : color.bg}; border-color: ${isCardHighlighted ? '#4F46E5' : (isCanceled ? '#CBD5E1' : color.border)};"
                     draggable="true"
                     ondragstart="window.app.handleTimetableCardDragStart(event, '${safeSub}', '${day.id}', '${period.id}')"
                     ondragend="window.app.handleTimetableCardDragEnd(event)">
                  
                  <!-- 카드 상단: 상태 기호 + 과목명 + 충돌 경고 + 삭제 버튼 -->
                  <div class="flex items-center justify-between gap-1 mb-1 pointer-events-auto">
                    <div class="flex items-center gap-1 min-w-0 flex-wrap">
                      ${isCardHighlighted ? '<span class="inline-flex items-center gap-0.5 text-[10px] font-black text-white bg-indigo-600 px-1.5 py-0.2 rounded shadow-2xs shrink-0 animate-pulse">✨ 강조</span>' : ''}
                      <!-- 상태 토글 버튼 (확정 -> 조정 -> 폐강) -->
                      <button type="button" 
                              onclick="event.stopPropagation(); window.app.toggleCourseScheduleStatus('${safeSub}', '${schedule.id || (day.id + '_' + period.id)}')"
                              class="hover:scale-110 transition shrink-0" title="클릭하여 상태 전환 (🟢확정 -> ▲조정 -> ✖️폐강)">
                        ${symHtml}
                      </button>
                      <span class="font-extrabold truncate" style="color: ${isCanceled ? '#64748B' : color.text};">
                        ${subject}
                      </span>
                      ${conflictCount > 0 ? `
                        <span class="inline-flex items-center gap-0.5 text-[11px] font-black text-rose-700 bg-rose-100 hover:bg-rose-200 border border-rose-300 px-1 py-0.2 rounded shadow-2xs cursor-help transition shrink-0" 
                              title="${conflictTooltip}">
                          <span>⚠️</span>
                          <span>${conflictCount}</span>
                        </span>
                      ` : ''}
                    </div>

                    <!-- 이 시간표 셀에서만 삭제 -->
                    <button type="button" 
                            onclick="event.stopPropagation(); window.app.handleRemoveCourseScheduleSlot('${safeSub}', '${schedule.id || (day.id + '_' + period.id)}')"
                            class="text-slate-400 hover:text-red-600 font-bold px-1 transition" title="이 시간 배치 취소">
                      &times;
                    </button>
                  </div>

                  <!-- 카드 하단: 학생 명단 칩 & 명단 팝업 버튼 -->
                  <div class="flex items-center justify-between text-[11px] pointer-events-auto">
                    <button type="button" 
                            onclick="event.stopPropagation(); window.app.openCourseRosterModal('${safeSub}', '${day.id}', '${period.id}')"
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
  // 시간표 카드 드래그 & 셀 드롭 (기본: 이동, Ctrl 누른 채 드래그 시: 복사)
  // ==========================================
  handleTimetableCardDragStart(e, subjectName, fromDay, fromPeriodId) {
    this.dragData = {
      type: 'timetable-subject',
      subjectName: subjectName,
      fromDay: fromDay,
      fromPeriodId: fromPeriodId
    };

    try {
      e.dataTransfer.setData('text/plain', JSON.stringify(this.dragData));
      e.dataTransfer.setData('application/json', JSON.stringify(this.dragData));
    } catch (err) {}
    e.dataTransfer.effectAllowed = 'copyMove';
    e.currentTarget.classList.add('is-dragging');

    if (e.ctrlKey || e.metaKey || this.isCtrlPressed) {
      const banner = document.getElementById('ctrl-copy-indicator');
      if (banner) banner.classList.remove('hidden');
    }
  }

  handleTimetableCardDragEnd(e) {
    this.dragData = null;
    if (e.currentTarget) e.currentTarget.classList.remove('is-dragging');
    document.querySelectorAll('.timetable-cell.drag-over').forEach(c => c.classList.remove('drag-over'));
    const banner = document.getElementById('ctrl-copy-indicator');
    if (banner && !e.ctrlKey && !e.metaKey) banner.classList.add('hidden');
  }

  handleTimetableCellDragOver(e) {
    e.preventDefault();
    const isPalette = this.dragData && this.dragData.type === 'palette-subject';
    // 시간표 카드는 오직 Ctrl/Meta 키가 실제로 눌려있을 때만 copy 효과 적용
    const isCopy = isPalette || Boolean(e.ctrlKey || e.metaKey);
    e.dataTransfer.dropEffect = isCopy ? 'copy' : 'move';
    const cell = e.currentTarget;
    if (cell && !cell.classList.contains('drag-over')) {
      cell.classList.add('drag-over');
    }
  }

  handleTimetableCellDragLeave(e) {
    e.currentTarget.classList.remove('drag-over');
  }

  handleTimetableCellDrop(e, targetDay, targetPeriodId) {
    e.preventDefault();
    if (e.currentTarget) e.currentTarget.classList.remove('drag-over');
    document.querySelectorAll('.timetable-cell.drag-over').forEach(c => c.classList.remove('drag-over'));

    let data = this.dragData;
    if (!data) {
      try {
        const raw = e.dataTransfer.getData('text/plain') || e.dataTransfer.getData('application/json');
        if (raw) data = JSON.parse(raw);
      } catch (err) {}
    }

    if (!data) return;

    const targetPeriod = this.periods.find(p => p.id === targetPeriodId);
    const dayName = ExcelService.dayIdToName(targetDay);
    const periodName = targetPeriod ? targetPeriod.name : `${targetPeriodId}교시`;

    // 1. 좌측 과목 팔레트에서 새 과목을 끌고 온 경우 (신규 배치)
    if (data.type === 'palette-subject') {
      const added = this.addCourseSchedule(data.subjectName, targetDay, targetPeriodId, 'confirmed');
      if (added) {
        const studentCount = (this.courseEnrollments[data.subjectName] || []).length;
        this.showToast(`'${data.subjectName}' 과목 (학생 ${studentCount}명)이 ${dayName} ${periodName}에 배치되었습니다.`);
      }
      this.dragData = null;
      return;
    }

    // 2. 시간표 내의 과목 카드를 다른 셀로 끌고 온 경우
    if (data.type === 'timetable-subject') {
      const { subjectName, fromDay, fromPeriodId } = data;

      // 같은 위치로 드롭한 경우 무시
      if (fromDay === targetDay && fromPeriodId === targetPeriodId) {
        this.dragData = null;
        return;
      }

      // 사용자가 드롭할 때 실제로 Ctrl 또는 Meta(Cmd) 키를 누르고 있는 경우에만 '복사', 그렇지 않으면 기본 '이동'
      const isCopyMode = Boolean(e.ctrlKey || e.metaKey);

      if (isCopyMode) {
        // [복사 (Copy)]: 기존 슬롯은 그대로 유지하고 대상 위치에 추가!
        const added = this.addCourseSchedule(subjectName, targetDay, targetPeriodId, 'confirmed');
        if (added) {
          this.showToast(`[Ctrl 복사] '${subjectName}' 과목이 ${dayName} ${periodName}에 추가 배치되었습니다.`);
        }
      } else {
        // [이동 (Move)]: 기존 위치 슬롯을 대상 위치로 옮김!
        const moved = this.moveCourseSchedule(subjectName, fromDay, fromPeriodId, targetDay, targetPeriodId);
        if (moved) {
          this.showToast(`'${subjectName}' 과목이 ${dayName} ${periodName}으로 이동되었습니다.`);
        }
      }
      this.dragData = null;
    }
  }

  addCourseSchedule(subjectName, day, periodId, status = 'confirmed') {
    if (!this.courseSchedules[subjectName]) this.courseSchedules[subjectName] = [];
    const periodObj = this.periods.find(p => p.id === periodId);

    // 중복 체크
    const exists = this.courseSchedules[subjectName].some(sch => 
      sch.day === day && (sch.periodId === periodId || (!sch.periodId && sch.period === (periodObj ? periodObj.period : null)))
    );
    if (exists) {
      this.showToast(`'${subjectName}' 과목은 이미 해당 교시에 배치되어 있습니다.`);
      return false;
    }

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
    return true;
  }

  moveCourseSchedule(subjectName, fromDay, fromPeriodId, toDay, toPeriodId) {
    if (!this.courseSchedules[subjectName]) return;
    const periodObj = this.periods.find(p => p.id === toPeriodId);

    // 대상 위치에 이미 동일 과목이 있는지 확인
    const exists = this.courseSchedules[subjectName].some(sch => 
      sch.day === toDay && (sch.periodId === toPeriodId || (!sch.periodId && sch.period === (periodObj ? periodObj.period : null)))
    );
    if (exists) {
      this.showToast(`'${subjectName}' 과목은 이미 해당 교시에 배치되어 있습니다.`);
      return false;
    }

    const slot = this.courseSchedules[subjectName].find(sch => 
      sch.day === fromDay && (sch.periodId === fromPeriodId || (!sch.periodId && sch.period === (this.periods.find(p => p.id === fromPeriodId)?.period)))
    );
    if (slot) {
      slot.day = toDay;
      slot.periodId = toPeriodId;
      slot.period = periodObj ? periodObj.period : null;
      this.rebuildStudentAssignments();
      this.saveState();
      this.render();
      return true;
    }
    return false;
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

        // 동일 시간(dayId, periodId)에 다른 유효 과목 수강 중인지 확인
        const conflictingSubjects = [];
        if (dayId && periodId) {
          Object.entries(this.courseSchedules).forEach(([otherSub, schList]) => {
            if (otherSub === subjectName) return;
            const hasSlot = schList.some(sch => {
              if (sch.status === 'canceled') return false;
              return sch.day === dayId && (sch.periodId === periodId || (!sch.periodId && sch.period === periodObj?.period));
            });
            if (hasSlot) {
              const otherEnrolled = this.courseEnrollments[otherSub] || [];
              if (otherEnrolled.includes(m.id)) {
                conflictingSubjects.push(otherSub);
              }
            }
          });
        }

        return `
          <div class="py-2.5 px-3 flex items-center justify-between text-xs hover:bg-slate-50 ${conflictingSubjects.length > 0 ? 'bg-rose-50/40' : ''}">
            <div class="flex items-center gap-2.5 flex-wrap">
              <span class="w-5 text-center text-slate-400 font-mono text-[11px]">${idx + 1}</span>
              <span class="font-mono font-bold text-slate-900">${code}</span>
              <span class="font-bold text-slate-800">${m.name}</span>
              ${conflictingSubjects.length > 0 ? `
                <span class="inline-flex items-center gap-1 text-[10px] font-bold text-rose-700 bg-rose-100/80 border border-rose-300 px-1.5 py-0.5 rounded shadow-2xs" 
                      title="동일 시간 타 과목 [${conflictingSubjects.join(', ')}] 동시 배정됨">
                  <span>⚠️ 중복:</span>
                  <span class="font-extrabold">${conflictingSubjects.join(', ')}</span>
                </span>
              ` : ''}
            </div>
            <button type="button" 
                    onclick="window.app.handleRemoveStudentFromCourse('${subjectName}', '${m.id}'); window.app.openCourseRosterModal('${subjectName}', '${dayId}', '${periodId}');" 
                    class="text-xs text-slate-400 hover:text-red-600 font-medium transition shrink-0 ml-2" title="수강 제외">
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

  // ==========================================
  // ==========================================
  // [학년도 및 학기 데이터 변경 / 복사 / 이동 매니저]
  // ==========================================
  openSemesterTransferModal(context = null) {
    const modal = document.getElementById('semester-transfer-modal');
    if (!modal) return;

    // 현재 출발지 뱃지
    const srcBadge = document.getElementById('transfer-source-badge');
    if (srcBadge) {
      srcBadge.textContent = `${this.activeYear}학년도 ${this.activeSemester}학기`;
    }

    // 도착 학년도 셀렉트 박스 채우기
    const yearSel = document.getElementById('transfer-target-year');
    if (yearSel) {
      yearSel.innerHTML = this.academicYears.map(y => {
        return `<option value="${y}" ${y === this.activeYear ? 'selected' : ''}>${y}학년도</option>`;
      }).join('');
    }

    // 도착 학기 셀렉트 박스: 기본값은 현재 학기의 반대 학기
    const semSel = document.getElementById('transfer-target-semester');
    if (semSel) {
      semSel.value = this.activeSemester === '1' ? '2' : '1';
    }

    // 과목 셀렉트 채우기
    const subSel = document.getElementById('transfer-subject-select');
    if (subSel) {
      subSel.innerHTML = this.offeredSubjects.map(s => {
        return `<option value="${s}">${s}</option>`;
      }).join('');
    }

    // 선택된 학생 수 표시
    const selectedBadge = document.getElementById('transfer-selected-count-badge');
    const selectedCount = this.selectedStep1StudentIds.size;
    if (selectedBadge) {
      selectedBadge.textContent = `${selectedCount}명 선택됨`;
    }

    // 라디오 버튼 초기화
    const radios = document.querySelectorAll('input[name="transfer-scope"]');
    if (context && typeof context === 'string' && this.offeredSubjects.includes(context)) {
      radios.forEach(r => { r.checked = r.value === 'subject'; });
      if (subSel) {
        subSel.value = context;
        subSel.disabled = false;
      }
    } else if (context === 'selected_students') {
      radios.forEach(r => { r.checked = r.value === 'selected_students'; });
      if (subSel) subSel.disabled = true;
    } else {
      radios.forEach(r => { r.checked = r.value === 'all'; });
      if (subSel) subSel.disabled = true;
    }

    modal.classList.remove('hidden');
    modal.classList.add('flex');
    if (window.lucide) window.lucide.createIcons();
  }

  closeSemesterTransferModal() {
    const modal = document.getElementById('semester-transfer-modal');
    if (!modal) return;
    modal.classList.add('hidden');
    modal.classList.remove('flex');
  }

  // 모달에서 특정 학년도·학기로 단순 화면 전환(변경)
  handleQuickSwitchSemesterFromModal() {
    const targetYear = document.getElementById('transfer-target-year')?.value || this.activeYear;
    const targetSemester = document.getElementById('transfer-target-semester')?.value || this.activeSemester;

    if (targetYear === this.activeYear && targetSemester === this.activeSemester) {
      this.closeSemesterTransferModal();
      this.showToast(`이미 현재 [${targetYear}학년도 ${targetSemester}학기] 화면입니다.`);
      return;
    }

    this.saveActiveYearSemesterData();
    if (!this.academicYears.includes(targetYear)) {
      this.academicYears.push(targetYear);
      this.academicYears.sort((a, b) => Number(a) - Number(b));
    }
    this.ensureYearSemesterData(targetYear, targetSemester);
    this.activeYear = targetYear;
    this.activeSemester = targetSemester;
    this.loadActiveYearSemesterData();
    this.sortPeriods();
    this.buildSubjectColorMap();
    this.saveState();
    this.render();
    this.closeSemesterTransferModal();
    this.showToast(`[${this.activeYear}학년도 ${this.activeSemester}학기] 시간표 화면으로 전환되었습니다.`);
  }

  handleSemesterTransferSubmit(e) {
    e.preventDefault();

    const scope = document.querySelector('input[name="transfer-scope"]:checked')?.value || 'all';
    const mode = document.querySelector('input[name="transfer-mode"]:checked')?.value || 'copy';
    const targetYear = document.getElementById('transfer-target-year')?.value || this.activeYear;
    const targetSemester = document.getElementById('transfer-target-semester')?.value || (this.activeSemester === '1' ? '2' : '1');
    const subjectName = document.getElementById('transfer-subject-select')?.value;

    // 출발지와 도착지가 완전히 동일한 경우 안내
    if (targetYear === this.activeYear && targetSemester === this.activeSemester) {
      alert('출발 학기와 도착 학기가 동일합니다. 데이터를 복사/이전할 다른 대상 학년도 또는 다른 학기를 선택해주세요.\n(화면만 변경하시려면 좌측의 [화면만 전환 (변경)] 버튼을 누르세요.)');
      return;
    }

    if (scope === 'subject' && !subjectName) {
      alert('전송할 과목을 선택해주세요.');
      return;
    }

    if (scope === 'selected_students' && this.selectedStep1StudentIds.size === 0) {
      alert('먼저 전송할 학생을 명단에서 체크 선택해주세요.');
      return;
    }

    // 1. 현재 학년도/학기 상태 저장 및 실행 취소 스냅샷 저장
    this.saveActiveYearSemesterData();
    this.pushUndoSnapshot(`학기 전송 (${this.activeYear}년 ${this.activeSemester}학기 -> ${targetYear}년 ${targetSemester}학기, ${mode === 'move' ? '이동' : '복사'})`);

    // 2. 대상 학년도 및 학기 데이터 준비
    if (!this.academicYears.includes(targetYear)) {
      this.academicYears.push(targetYear);
      this.academicYears.sort((a, b) => Number(a) - Number(b));
    }
    this.ensureYearSemesterData(targetYear, targetSemester);
    const targetData = this.yearsData[targetYear][targetSemester];

    // Helper: 대상 학기에 학생 중복 없이 추가
    const addStudentToTarget = (st) => {
      const pCode = st.studentCode || ExcelService.getStudentCode(st);
      const exists = (targetData.masterStudents || []).find(ms => {
        const mCode = ms.studentCode || ExcelService.getStudentCode(ms);
        return mCode === pCode && ms.name === st.name;
      });
      if (exists) return exists.id;

      const newId = `sm_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`;
      const newObj = {
        id: newId,
        grade: st.grade,
        classNum: st.classNum,
        studentNum: st.studentNum,
        studentCode: pCode,
        name: st.name
      };
      if (!targetData.masterStudents) targetData.masterStudents = [];
      targetData.masterStudents.push(newObj);
      return newId;
    };

    let summaryMsg = '';

    if (scope === 'all') {
      // 1) 전체 데이터 복사 / 이동
      targetData.periods = JSON.parse(JSON.stringify(this.periods));
      targetData.days = JSON.parse(JSON.stringify(this.days));
      targetData.students = JSON.parse(JSON.stringify(this.students));
      targetData.masterStudents = JSON.parse(JSON.stringify(this.masterStudents || []));
      targetData.courseEnrollments = JSON.parse(JSON.stringify(this.courseEnrollments || {}));
      targetData.courseSchedules = JSON.parse(JSON.stringify(this.courseSchedules || {}));
      targetData.courseStatus = JSON.parse(JSON.stringify(this.courseStatus || {}));
      targetData.offeredSubjects = [...this.offeredSubjects];
      targetData.subjectColorMap = { ...this.subjectColorMap };
      targetData.showWeekends = this.showWeekends;

      if (mode === 'move') {
        // 현재 학기는 초기화
        this.masterStudents = [];
        this.courseEnrollments = {};
        this.courseSchedules = {};
        this.courseStatus = {};
        this.students = [];
        this.offeredSubjects = [...DEFAULT_OFFERED_SUBJECTS];
        this.buildSubjectColorMap();
        this.rebuildStudentAssignments();
        this.saveActiveYearSemesterData();
      }

      summaryMsg = `현재 학기 전체 데이터가 [${targetYear}학년도 ${targetSemester}학기]로 ${mode === 'move' ? '이동' : '복사'}되었습니다.`;
    } 
    else if (scope === 'subject') {
      // 2) 특정 과목 전송 (수강생, 시간표 및 개설/조정/폐강 상태 포함)
      if (!targetData.offeredSubjects.includes(subjectName)) {
        targetData.offeredSubjects.push(subjectName);
      }
      if (!targetData.subjectColorMap[subjectName] && this.subjectColorMap[subjectName]) {
        targetData.subjectColorMap[subjectName] = this.subjectColorMap[subjectName];
      }
      if (!targetData.courseStatus) targetData.courseStatus = {};
      targetData.courseStatus[subjectName] = this.getSubjectStatus(subjectName);

      // 수강생 매핑 및 이전 (기존 수강생과 안전 병합)
      const enrolledStudentIds = this.courseEnrollments[subjectName] || [];
      const targetStudentIds = [];
      enrolledStudentIds.forEach(stId => {
        const stObj = this.masterStudents.find(s => s.id === stId);
        if (stObj) {
          const targetId = addStudentToTarget(stObj);
          targetStudentIds.push(targetId);
        }
      });
      if (!targetData.courseEnrollments) targetData.courseEnrollments = {};
      const existingEnrollments = targetData.courseEnrollments[subjectName] || [];
      targetStudentIds.forEach(id => {
        if (!existingEnrollments.includes(id)) existingEnrollments.push(id);
      });
      targetData.courseEnrollments[subjectName] = existingEnrollments;

      // 시간표 스케줄 이전
      if (!targetData.courseSchedules) targetData.courseSchedules = {};
      targetData.courseSchedules[subjectName] = JSON.parse(JSON.stringify(this.courseSchedules[subjectName] || []));

      if (mode === 'move') {
        this.offeredSubjects = this.offeredSubjects.filter(s => s !== subjectName);
        delete this.courseEnrollments[subjectName];
        delete this.courseSchedules[subjectName];
        if (this.courseStatus) delete this.courseStatus[subjectName];
        this.rebuildStudentAssignments();
        this.saveActiveYearSemesterData();
      }

      summaryMsg = `'${subjectName}' 과목이 [${targetYear}학년도 ${targetSemester}학기]로 ${mode === 'move' ? '이동' : '복사'}되었습니다.`;
    }
    else if (scope === 'students') {
      // 3) 전체 학생 명단만 전송 (과목 편제는 제외)
      let count = 0;
      this.masterStudents.forEach(st => {
        addStudentToTarget(st);
        count++;
      });

      if (mode === 'move') {
        this.masterStudents = [];
        this.courseEnrollments = {};
        this.courseSchedules = {};
        this.students = [];
        this.rebuildStudentAssignments();
        this.saveActiveYearSemesterData();
      }

      summaryMsg = `학생 명단 ${count}명이 [${targetYear}학년도 ${targetSemester}학기]로 ${mode === 'move' ? '이동' : '복사'}되었습니다.`;
    }
    else if (scope === 'selected_students') {
      // 4) 선택된 학생들만 전송
      let count = 0;
      this.selectedStep1StudentIds.forEach(stId => {
        const stObj = this.masterStudents.find(s => s.id === stId);
        if (stObj) {
          addStudentToTarget(stObj);
          count++;
        }
      });

      if (mode === 'move') {
        this.masterStudents = this.masterStudents.filter(s => !this.selectedStep1StudentIds.has(s.id));
        Object.keys(this.courseEnrollments).forEach(sub => {
          this.courseEnrollments[sub] = this.courseEnrollments[sub].filter(id => !this.selectedStep1StudentIds.has(id));
        });
        this.selectedStep1StudentIds.clear();
        this.rebuildStudentAssignments();
        this.saveActiveYearSemesterData();
      }

      summaryMsg = `선택한 학생 ${count}명이 [${targetYear}학년도 ${targetSemester}학기]로 ${mode === 'move' ? '이동' : '복사'}되었습니다.`;
    }

    // 대상 학기의 assignments 재생성 및 영구 저장
    this.rebuildSemesterAssignments(targetData);
    this.saveState();
    this.render();
    this.closeSemesterTransferModal();

    // 완료 알림 및 해당 학기로 즉시 이동 여부 질문
    const moveConfirm = confirm(`${summaryMsg}\n\n[확인]: 지금 바로 [${targetYear}학년도 ${targetSemester}학기] 시간표로 이동\n[취소]: 현재 화면 유지`);
    if (moveConfirm) {
      this.activeYear = targetYear;
      this.activeSemester = targetSemester;
      this.loadActiveYearSemesterData();
      this.sortPeriods();
      this.buildSubjectColorMap();
      this.saveState();
      this.render();
      this.showToast(`[${this.activeYear}학년도 ${this.activeSemester}학기] 화면으로 전환되었습니다.`, true);
    } else {
      this.showToast(summaryMsg, true);
    }
  }

  rebuildSemesterAssignments(semData) {
    if (!semData) return;
    const flat = [];
    const masterMap = new Map((semData.masterStudents || []).map(s => [s.id, s]));
    const subjects = semData.offeredSubjects || [];

    subjects.forEach(subject => {
      const studentIds = (semData.courseEnrollments || {})[subject] || [];
      const schedules = (semData.courseSchedules || {})[subject] || [];

      studentIds.forEach(stId => {
        const m = masterMap.get(stId);
        if (!m) return;
        const code = m.studentCode || ExcelService.getStudentCode(m);

        if (schedules.length === 0) {
          flat.push({
            id: `as_${m.id}_${subject}_none`,
            studentCode: code,
            name: m.name,
            grade: m.grade,
            classNum: m.classNum,
            studentNum: m.studentNum,
            subject: subject,
            day: '',
            period: '',
            periodId: '',
            status: 'confirmed'
          });
        } else {
          schedules.forEach(sch => {
            flat.push({
              id: `as_${m.id}_${subject}_${sch.id || (sch.day + '_' + sch.periodId)}`,
              studentCode: code,
              name: m.name,
              grade: m.grade,
              classNum: m.classNum,
              studentNum: m.studentNum,
              subject: subject,
              day: sch.day,
              period: sch.period,
              periodId: sch.periodId,
              status: sch.status || 'confirmed'
            });
          });
        }
      });
    });

    semData.students = flat;
  }

  // ==========================================
  // [1. 과목명 수정 (Rename Subject)]
  // ==========================================
  handleRenameSubject(oldName) {
    if (!oldName) return;
    const input = prompt(`'${oldName}' 과목의 새 이름을 입력하세요:`, oldName);
    if (input === null) return;
    const newName = input.trim();
    if (!newName) {
      alert('과목명을 공백으로 둘 수 없습니다.');
      return;
    }
    if (newName === oldName) return;

    // 이미 존재하는 과목명인지 검사 (중복 방지 경고 정책)
    if (this.offeredSubjects.includes(newName)) {
      alert(`'${newName}' 과목은 이미 개설 과목 목록에 존재합니다.\n다른 과목명을 입력해주세요.`);
      return;
    }

    // 1) offeredSubjects 배열 내 변경 (순서 유지)
    const idx = this.offeredSubjects.indexOf(oldName);
    if (idx !== -1) {
      this.offeredSubjects[idx] = newName;
    }

    // 2) courseEnrollments 변경
    if (this.courseEnrollments[oldName]) {
      this.courseEnrollments[newName] = this.courseEnrollments[oldName];
      delete this.courseEnrollments[oldName];
    } else {
      this.courseEnrollments[newName] = [];
    }

    // 3) courseSchedules 변경
    if (this.courseSchedules[oldName]) {
      this.courseSchedules[newName] = this.courseSchedules[oldName];
      delete this.courseSchedules[oldName];
    }

    // 4) courseStatus 변경
    if (this.courseStatus && this.courseStatus[oldName]) {
      this.courseStatus[newName] = this.courseStatus[oldName];
      delete this.courseStatus[oldName];
    }

    // 5) subjectColorMap 변경
    if (this.subjectColorMap[oldName]) {
      this.subjectColorMap[newName] = this.subjectColorMap[oldName];
      delete this.subjectColorMap[oldName];
    } else {
      this.buildSubjectColorMap();
    }

    // 6) flat students 변경
    if (this.students && Array.isArray(this.students)) {
      this.students.forEach(s => {
        if (s.subject === oldName) {
          s.subject = newName;
        }
      });
    }

    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.showToast(`과목명이 '${oldName}'에서 '${newName}'(으)로 변경되었습니다.`);
  }

  // ==========================================
  // [2. 과목 전용 학기 이전 / 복사 매니저]
  // ==========================================
  openSubjectTransferModal(subject) {
    const modal = document.getElementById('subject-transfer-modal');
    if (!modal) return;

    const subNameBadge = document.getElementById('transfer-sub-name-badge');
    const subHiddenName = document.getElementById('transfer-sub-hidden-name');
    const srcBadge = document.getElementById('transfer-sub-source-badge');

    if (subNameBadge) subNameBadge.textContent = subject;
    if (subHiddenName) subHiddenName.value = subject;
    if (srcBadge) srcBadge.textContent = `${this.activeYear}학년도 ${this.activeSemester}학기`;

    // 도착 학년도 셀렉트 박스 채우기
    const yearSel = document.getElementById('transfer-sub-target-year');
    if (yearSel) {
      let options = this.academicYears.map(y => {
        return `<option value="${y}" ${y === this.activeYear ? 'selected' : ''}>${y}학년도</option>`;
      });
      options.push('<option value="custom">+ 새 학년도 직접 입력...</option>');
      yearSel.innerHTML = options.join('');
    }

    const customWrap = document.getElementById('transfer-sub-custom-year-wrap');
    if (customWrap) customWrap.classList.add('hidden');
    const customInput = document.getElementById('transfer-sub-custom-year');
    if (customInput) customInput.value = '';

    // 도착 학기 기본값: 반대 학기
    const targetSemRadios = document.querySelectorAll('input[name="transfer-sub-target-semester"]');
    const oppositeSem = this.activeSemester === '1' ? '2' : '1';
    targetSemRadios.forEach(r => {
      r.checked = r.value === oppositeSem;
    });

    // 작업 방식 기본값: 복사
    const modeRadios = document.querySelectorAll('input[name="transfer-sub-mode"]');
    modeRadios.forEach(r => {
      r.checked = r.value === 'copy';
    });

    modal.classList.remove('hidden');
    modal.classList.add('flex');
    if (window.lucide) window.lucide.createIcons();
  }

  closeSubjectTransferModal() {
    const modal = document.getElementById('subject-transfer-modal');
    if (!modal) return;
    modal.classList.add('hidden');
    modal.classList.remove('flex');
  }

  handleTransferSubYearChange(val) {
    const customWrap = document.getElementById('transfer-sub-custom-year-wrap');
    const customInput = document.getElementById('transfer-sub-custom-year');
    if (!customWrap) return;

    if (val === 'custom') {
      customWrap.classList.remove('hidden');
      if (customInput) {
        customInput.focus();
        if (!customInput.value) {
          const nextYear = String(Math.max(...this.academicYears.map(Number)) + 1);
          customInput.value = nextYear;
        }
      }
    } else {
      customWrap.classList.add('hidden');
    }
  }

  handleSubjectTransferSubmit(e) {
    e.preventDefault();

    const subjectName = document.getElementById('transfer-sub-hidden-name')?.value;
    if (!subjectName) {
      alert('대상 과목 정보가 올바르지 않습니다.');
      return;
    }

    const yearSelVal = document.getElementById('transfer-sub-target-year')?.value || this.activeYear;
    let targetYear = yearSelVal;
    if (yearSelVal === 'custom') {
      const customYearVal = document.getElementById('transfer-sub-custom-year')?.value.trim().replace(/[^0-9]/g, '');
      if (!customYearVal || customYearVal.length !== 4) {
        alert('이동/복사할 대상 학년도를 4자리 숫자로 입력해주세요 (예: 2028).');
        document.getElementById('transfer-sub-custom-year')?.focus();
        return;
      }
      targetYear = customYearVal;
    }

    const targetSemester = document.querySelector('input[name="transfer-sub-target-semester"]:checked')?.value || (this.activeSemester === '1' ? '2' : '1');
    const mode = document.querySelector('input[name="transfer-sub-mode"]:checked')?.value || 'copy';

    // 출발지와 도착지가 완전히 동일한 경우 방지
    if (targetYear === this.activeYear && targetSemester === this.activeSemester) {
      alert('출발 학기와 도착 학기가 동일합니다. 다른 학년도 또는 다른 학기를 선택해주세요.');
      return;
    }

    // 1. 현재 학년도/학기 상태 저장 및 실행 취소 스냅샷 저장
    this.saveActiveYearSemesterData();
    this.pushUndoSnapshot(`과목 '${subjectName}' ${mode === 'move' ? '이동' : '복사'} (${this.activeYear}년 ${this.activeSemester}학기 -> ${targetYear}년 ${targetSemester}학기)`);

    // 2. 대상 학년도 및 학기 데이터 준비
    if (!this.academicYears.includes(targetYear)) {
      this.academicYears.push(targetYear);
      this.academicYears.sort((a, b) => Number(a) - Number(b));
    }
    this.ensureYearSemesterData(targetYear, targetSemester);
    const targetData = this.yearsData[targetYear][targetSemester];

    // Helper: 대상 학기에 학생 중복 없이 추가
    const addStudentToTarget = (st) => {
      const pCode = st.studentCode || ExcelService.getStudentCode(st);
      const exists = (targetData.masterStudents || []).find(ms => {
        const mCode = ms.studentCode || ExcelService.getStudentCode(ms);
        return mCode === pCode && ms.name === st.name;
      });
      if (exists) return exists.id;

      const newId = `sm_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`;
      const newObj = {
        id: newId,
        grade: st.grade,
        classNum: st.classNum,
        studentNum: st.studentNum,
        studentCode: pCode,
        name: st.name
      };
      if (!targetData.masterStudents) targetData.masterStudents = [];
      targetData.masterStudents.push(newObj);
      return newId;
    };

    // 3. 과목 등록 및 메타데이터 복사
    if (!targetData.offeredSubjects.includes(subjectName)) {
      targetData.offeredSubjects.push(subjectName);
    }
    if (!targetData.subjectColorMap[subjectName] && this.subjectColorMap[subjectName]) {
      targetData.subjectColorMap[subjectName] = this.subjectColorMap[subjectName];
    }
    if (!targetData.courseStatus) targetData.courseStatus = {};
    targetData.courseStatus[subjectName] = this.getSubjectStatus(subjectName);

    // 4. 수강생 매핑 및 이전 (기존 수강생과 안전 병합)
    const enrolledStudentIds = this.courseEnrollments[subjectName] || [];
    const targetStudentIds = [];
    enrolledStudentIds.forEach(stId => {
      const stObj = this.masterStudents.find(s => s.id === stId);
      if (stObj) {
        const targetId = addStudentToTarget(stObj);
        targetStudentIds.push(targetId);
      }
    });
    if (!targetData.courseEnrollments) targetData.courseEnrollments = {};
    const existingEnrollments = targetData.courseEnrollments[subjectName] || [];
    targetStudentIds.forEach(id => {
      if (!existingEnrollments.includes(id)) existingEnrollments.push(id);
    });
    targetData.courseEnrollments[subjectName] = existingEnrollments;

    // 5. 시간표 스케줄 이전
    if (!targetData.courseSchedules) targetData.courseSchedules = {};
    targetData.courseSchedules[subjectName] = JSON.parse(JSON.stringify(this.courseSchedules[subjectName] || []));

    // 6. 이동 모드인 경우 현재 학기에서 삭제
    if (mode === 'move') {
      this.offeredSubjects = this.offeredSubjects.filter(s => s !== subjectName);
      delete this.courseEnrollments[subjectName];
      delete this.courseSchedules[subjectName];
      if (this.courseStatus) delete this.courseStatus[subjectName];
      this.rebuildStudentAssignments();
      this.saveActiveYearSemesterData();
    }

    // 대상 학기의 assignments 재생성 및 영구 저장
    this.rebuildSemesterAssignments(targetData);
    this.saveState();
    this.render();
    this.closeSubjectTransferModal();

    const summaryMsg = `'${subjectName}' 과목이 [${targetYear}학년도 ${targetSemester}학기]로 ${mode === 'move' ? '이동' : '복사'}되었습니다.`;
    const moveConfirm = confirm(`${summaryMsg}\n\n[확인]: 지금 바로 [${targetYear}학년도 ${targetSemester}학기] 시간표로 이동\n[취소]: 현재 화면 유지`);
    if (moveConfirm) {
      this.activeYear = targetYear;
      this.activeSemester = targetSemester;
      this.loadActiveYearSemesterData();
      this.sortPeriods();
      this.buildSubjectColorMap();
      this.saveState();
      this.render();
      this.showToast(`[${this.activeYear}학년도 ${this.activeSemester}학기] 화면으로 전환되었습니다.`, true);
    } else {
      this.showToast(summaryMsg, true);
    }
  }

  // ==========================================
  // [3. 학생 정보 및 수강 과목 수정 매니저]
  // ==========================================
  openEditStudentModal(studentId) {
    const student = this.masterStudents.find(s => s.id === studentId);
    if (!student) {
      alert('학생 정보를 찾을 수 없습니다.');
      return;
    }

    const modal = document.getElementById('student-edit-modal');
    if (!modal) return;

    document.getElementById('edit-student-id').value = student.id;
    document.getElementById('edit-student-grade').value = student.grade || 1;
    document.getElementById('edit-student-class').value = student.classNum || 1;
    document.getElementById('edit-student-num').value = student.studentNum !== null && student.studentNum !== undefined ? student.studentNum : 1;
    document.getElementById('edit-student-name').value = student.name || '';
    document.getElementById('edit-student-code').value = student.studentCode || ExcelService.getStudentCode(student);

    // 수강 중인 과목 체크박스 렌더링
    const container = document.getElementById('edit-student-subject-checkboxes');
    const countBadge = document.getElementById('edit-student-enrollment-count');

    let enrolledCount = 0;
    if (container) {
      if (this.offeredSubjects.length === 0) {
        container.innerHTML = `<div class="col-span-full py-3 text-center text-slate-400 italic">개설된 과목 바구니가 없습니다.</div>`;
      } else {
        container.innerHTML = this.offeredSubjects.map(sub => {
          const isEnrolled = (this.courseEnrollments[sub] || []).includes(studentId);
          if (isEnrolled) enrolledCount++;
          const color = this.getSubjectColor(sub);
          const safeSub = encodeURIComponent(sub);

          return `
            <label class="p-2 rounded-lg border cursor-pointer transition flex items-center gap-2 ${
              isEnrolled ? 'bg-indigo-50/70 border-indigo-300' : 'bg-white border-slate-200 hover:bg-slate-50'
            }">
              <input type="checkbox" name="edit-student-course" value="${safeSub}" ${isEnrolled ? 'checked' : ''}
                     onchange="window.app.updateEditStudentEnrolledCount()"
                     class="text-indigo-600 focus:ring-indigo-500 rounded">
              <span class="w-2.5 h-2.5 rounded-full shrink-0" style="background-color: ${color.text};"></span>
              <span class="font-bold text-slate-800 truncate">${sub}</span>
            </label>
          `;
        }).join('');
      }
    }

    if (countBadge) {
      countBadge.textContent = `${enrolledCount}과목 수강 중`;
    }

    modal.classList.remove('hidden');
    modal.classList.add('flex');
    if (window.lucide) window.lucide.createIcons();
  }

  updateEditStudentEnrolledCount() {
    const checked = document.querySelectorAll('input[name="edit-student-course"]:checked');
    const badge = document.getElementById('edit-student-enrollment-count');
    if (badge) {
      badge.textContent = `${checked.length}과목 수강 중`;
    }
  }

  autoRecalculateEditStudentCode() {
    const grade = parseInt(document.getElementById('edit-student-grade')?.value, 10) || 1;
    const classNum = parseInt(document.getElementById('edit-student-class')?.value, 10) || 1;
    const studentNum = parseInt(document.getElementById('edit-student-num')?.value, 10) || 1;
    const code = ExcelService.getStudentCode({ grade, classNum, studentNum });
    const codeInput = document.getElementById('edit-student-code');
    if (codeInput) codeInput.value = code;
  }

  closeEditStudentModal() {
    const modal = document.getElementById('student-edit-modal');
    if (!modal) return;
    modal.classList.add('hidden');
    modal.classList.remove('flex');
  }

  handleEditStudentSubmit(e) {
    e.preventDefault();

    const studentId = document.getElementById('edit-student-id')?.value;
    const student = this.masterStudents.find(s => s.id === studentId);
    if (!student) {
      alert('학생 정보를 찾을 수 없습니다.');
      return;
    }

    const grade = parseInt(document.getElementById('edit-student-grade').value, 10) || 1;
    const classNum = parseInt(document.getElementById('edit-student-class').value, 10) || 1;
    const studentNum = parseInt(document.getElementById('edit-student-num').value, 10) || 1;
    const name = document.getElementById('edit-student-name').value.trim();
    const codeInput = document.getElementById('edit-student-code').value.trim();
    const studentCode = codeInput || ExcelService.getStudentCode({ grade, classNum, studentNum });

    if (!name) {
      alert('학생 성명을 입력해주세요.');
      return;
    }

    // 인적사항 갱신
    student.grade = grade;
    student.classNum = classNum;
    student.studentNum = studentNum;
    student.name = name;
    student.studentCode = studentCode;

    // 수강 과목 갱신
    const checkedInputs = Array.from(document.querySelectorAll('input[name="edit-student-course"]:checked'));
    const selectedSubjects = new Set(checkedInputs.map(input => decodeURIComponent(input.value)));

    this.offeredSubjects.forEach(sub => {
      if (!this.courseEnrollments[sub]) this.courseEnrollments[sub] = [];
      const hasStudent = this.courseEnrollments[sub].includes(studentId);
      const shouldHave = selectedSubjects.has(sub);

      if (shouldHave && !hasStudent) {
        this.courseEnrollments[sub].push(studentId);
      } else if (!shouldHave && hasStudent) {
        this.courseEnrollments[sub] = this.courseEnrollments[sub].filter(id => id !== studentId);
      }
    });

    this.rebuildStudentAssignments();
    this.saveState();
    this.render();
    this.closeEditStudentModal();
    this.showToast(`[${studentCode} ${name}] 학생 정보 및 수강 과목이 수정되었습니다.`);
  }

  handleDeleteStudentFromModal() {
    const studentId = document.getElementById('edit-student-id')?.value;
    const student = this.masterStudents.find(s => s.id === studentId);
    if (!student) return;

    if (!confirm(`'${student.name}' 학생을 명단 및 모든 수강 과목에서 완전히 삭제하시겠습니까?`)) {
      return;
    }

    this.deleteMasterStudent(studentId, true);
    this.closeEditStudentModal();
  }

  // ==========================================
  // [4. 학년도 관리 모달 매니저]
  // ==========================================
  openAcademicYearModal() {
    const modal = document.getElementById('academic-year-modal');
    if (!modal) return;

    this.renderAcademicYearListModal();
    const input = document.getElementById('input-new-academic-year');
    if (input) {
      input.value = String(Math.max(...this.academicYears.map(Number)) + 1);
    }

    modal.classList.remove('hidden');
    modal.classList.add('flex');
    if (window.lucide) window.lucide.createIcons();
  }

  closeAcademicYearModal() {
    const modal = document.getElementById('academic-year-modal');
    if (!modal) return;
    modal.classList.add('hidden');
    modal.classList.remove('flex');
  }

  renderAcademicYearListModal() {
    const container = document.getElementById('academic-year-list-container');
    if (!container) return;

    container.innerHTML = this.academicYears.map(year => {
      const isActive = year === this.activeYear;
      const yearObj = this.yearsData[year] || {};
      const count1 = (yearObj['1'] && yearObj['1'].masterStudents) ? yearObj['1'].masterStudents.length : 0;
      const count2 = (yearObj['2'] && yearObj['2'].masterStudents) ? yearObj['2'].masterStudents.length : 0;

      return `
        <div class="p-3 rounded-xl border flex items-center justify-between gap-3 ${
          isActive ? 'bg-blue-50/70 border-blue-300 ring-1 ring-blue-400/30' : 'bg-white border-slate-200 hover:bg-slate-50'
        }">
          <div class="flex items-center gap-2.5">
            <span class="w-2.5 h-2.5 rounded-full ${isActive ? 'bg-blue-600' : 'bg-slate-300'}"></span>
            <div>
              <div class="flex items-center gap-1.5">
                <span class="font-black text-sm text-slate-800">${year}학년도</span>
                ${isActive ? '<span class="text-[10px] font-bold px-1.5 py-0.2 rounded bg-blue-600 text-white">현재 활성</span>' : ''}
              </div>
              <p class="text-[11px] text-slate-500 mt-0.5">
                1학기: <strong>${count1}명</strong> · 2학기: <strong>${count2}명</strong> 등록됨
              </p>
            </div>
          </div>

          <div class="flex items-center gap-1.5">
            ${!isActive ? `
              <button type="button" 
                      onclick="window.app.switchAcademicYear('${year}'); window.app.renderAcademicYearListModal();" 
                      class="px-2.5 py-1 text-xs font-bold rounded-lg bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 transition" 
                      title="${year}학년도로 화면 전환">
                전환
              </button>
            ` : ''}

            ${this.academicYears.length > 1 ? `
              <button type="button" 
                      onclick="window.app.handleDeleteAcademicYear('${year}'); window.app.renderAcademicYearListModal();" 
                      class="px-2 py-1 text-xs font-bold rounded-lg hover:bg-rose-50 text-slate-400 hover:text-rose-600 transition" 
                      title="'${year}학년도' 삭제">
                <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
              </button>
            ` : ''}
          </div>
        </div>
      `;
    }).join('');

    if (window.lucide) window.lucide.createIcons();
  }

  handleCreateNewAcademicYearFromModal() {
    const input = document.getElementById('input-new-academic-year');
    if (!input) return;

    const val = input.value.trim().replace(/[^0-9]/g, '');
    if (!val || val.length !== 4) {
      alert('올바른 4자리 연도를 입력해주세요 (예: 2028)');
      input.focus();
      return;
    }

    if (this.academicYears.includes(val)) {
      alert(`'${val}학년도'는 이미 등록되어 있습니다.`);
      return;
    }

    this.saveActiveYearSemesterData();
    this.academicYears.push(val);
    this.academicYears.sort((a, b) => Number(a) - Number(b));
    this.yearsData[val] = {
      '1': this.createDefaultSemesterData(this.periods, this.days, this.offeredSubjects, this.subjectColorMap),
      '2': this.createDefaultSemesterData(this.periods, this.days, this.offeredSubjects, this.subjectColorMap)
    };
    this.activeYear = val;
    this.activeSemester = '1';
    this.loadActiveYearSemesterData();
    this.saveState();
    this.render();
    this.renderAcademicYearListModal();
    input.value = String(Number(val) + 1);
    this.showToast(`'${val}학년도'가 성공적으로 추가되었습니다.`);
  }

  showToast(message, allowUndo = false) {
    let toast = document.getElementById('app-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'app-toast';
      toast.className = 'fixed bottom-5 right-5 z-50 bg-slate-900 text-white text-xs px-4 py-3 rounded-xl shadow-2xl flex items-center gap-3.5 transform transition-all duration-300 opacity-0 translate-y-3 border border-slate-700 max-w-lg';
      document.body.appendChild(toast);
    }

    const hasUndo = allowUndo && this.undoStack && this.undoStack.length > 0;
    toast.innerHTML = `
      <div class="flex items-center gap-2 min-w-0">
        <i data-lucide="check-circle" class="w-4 h-4 text-emerald-400 shrink-0"></i>
        <span class="font-medium leading-relaxed">${message}</span>
      </div>
      ${hasUndo ? `
        <button type="button" 
                onclick="window.app.handleUndo()" 
                class="px-2.5 py-1 text-xs font-bold rounded-lg bg-amber-500 hover:bg-amber-600 active:scale-95 text-white transition flex items-center gap-1 shadow-xs shrink-0 cursor-pointer whitespace-nowrap ml-1"
                title="방금 수행한 작업을 즉시 취소하고 되돌립니다 (Ctrl+Z)">
          <i data-lucide="undo-2" class="w-3.5 h-3.5"></i>
          <span>실행 취소(되돌리기)</span>
        </button>
      ` : ''}
    `;
    if (window.lucide) window.lucide.createIcons();

    toast.classList.remove('opacity-0', 'translate-y-3');
    toast.classList.add('opacity-100', 'translate-y-0');

    if (this.toastTimeout) clearTimeout(this.toastTimeout);
    const duration = hasUndo ? 6500 : 3000;
    this.toastTimeout = setTimeout(() => {
      toast.classList.remove('opacity-100', 'translate-y-0');
      toast.classList.add('opacity-0', 'translate-y-3');
    }, duration);
  }
}

document.addEventListener('DOMContentLoaded', () => {
  window.app = new TimetableApp();
});
