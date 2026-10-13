/**
 * Excel Service using ExcelJS for rich styling, pastel colors, and multi-sheet support.
 * - 학생 편성 엑셀 업로드 시 [학번, 성명]만 추출
 * - 3단 상태(🟢 확정 / 🔺 조정 / ✖️ 폐강 - 회색) 서식 지원
 * - 파스텔톤 멀티시트 학급 시간표 및 원천 데이터 내보내기
 */

const ExcelService = {
  dayNameToId(name) {
    if (!name) return null;
    const clean = String(name).trim();
    const map = {
      '월': 'mon', '월요일': 'mon', 'mon': 'mon',
      '화': 'tue', '화요일': 'tue', 'tue': 'tue',
      '수': 'wed', '수요일': 'wed', 'wed': 'wed',
      '목': 'thu', '목요일': 'thu', 'thu': 'thu',
      '금': 'fri', '금요일': 'fri', 'fri': 'fri',
      '토': 'sat', '토요일': 'sat', 'sat': 'sat',
      '일': 'sun', '일요일': 'sun', 'sun': 'sun'
    };
    return map[clean] || null;
  },

  dayIdToName(id) {
    const map = { mon: '월', tue: '화', wed: '수', thu: '목', fri: '금', sat: '토', sun: '일' };
    return map[id] || '';
  },

  // 5자리 학번 생성: 1학년 1반 1번 -> 10101
  getStudentCode(student) {
    const g = student.grade ? String(student.grade).trim() : '1';
    const c = student.classNum !== undefined && student.classNum !== null ? String(student.classNum).trim().padStart(2, '0') : '01';
    if (student.studentNum !== undefined && student.studentNum !== null && String(student.studentNum).trim() !== '') {
      const n = String(student.studentNum).trim().padStart(2, '0');
      return `${g}${c}${n}`;
    }
    return `${g}${c}`;
  },

  // 학번이름: 10101홍길동
  getStudentCodeName(student) {
    const code = this.getStudentCode(student);
    return `${code}${student.name || ''}`;
  },

  // 상태별 기호 (확정: 🟢, 조정: 노란 세모 ▲, 폐강: ✖️)
  getStatusSymbol(status) {
    if (status === 'adjust') return '▲';
    if (status === 'canceled') return '✖️';
    return '🟢';
  },

  // 상태별 HTML 표기 (노란색 세모 스타일 적용)
  getStatusHtml(status) {
    if (status === 'adjust') {
      return '<span class="text-amber-500 font-black drop-shadow-xs" style="color: #F59E0B;" title="조정(노랑 세모)">▲</span>';
    }
    if (status === 'canceled') {
      return '<span class="text-slate-400 font-bold" title="폐강(회색 엑스)">✖️</span>';
    }
    return '<span class="text-emerald-500" title="확정(초록)">🟢</span>';
  },

  // 상태 텍스트
  getStatusText(status) {
    if (status === 'adjust') return '조정(▲)';
    if (status === 'canceled') return '폐강(✖️)';
    return '확정(🟢)';
  },

  // 상태 문자열 표준화 (개설/확정, 조정, 폐강)
  normalizeStatus(val) {
    if (!val) return { courseStatus: 'open', scheduleStatus: 'confirmed', label: '확정' };
    const str = String(val).trim().toLowerCase();
    if (str.includes('조정') || str.includes('▲') || str.includes('adjust')) {
      return { courseStatus: 'adjust', scheduleStatus: 'adjust', label: '조정' };
    }
    if (str.includes('폐강') || str.includes('✖') || str.includes('x') || str.includes('closed') || str.includes('cancel')) {
      return { courseStatus: 'closed', scheduleStatus: 'canceled', label: '폐강' };
    }
    return { courseStatus: 'open', scheduleStatus: 'confirmed', label: '확정' };
  },

  // 시간표 카드 표기: 과목명(학번성명)
  getTimetableCardLabel(student) {
    const sub = student.subject || '과목';
    const codeName = this.getStudentCodeName(student);
    const sym = this.getStatusSymbol(student.status);
    return `${sym} ${sub}(${codeName})`;
  },

  // 1. 일과표 양식 다운로드
  async downloadPeriodTemplate() {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = '강원온라인학교 시간표 시스템';
    const worksheet = workbook.addWorksheet('일과표 설정');

    worksheet.columns = [
      { header: '교시구분', key: 'type', width: 14 },
      { header: '교시번호', key: 'period', width: 12 },
      { header: '교시명', key: 'name', width: 16 },
      { header: '시작시간(HH:mm)', key: 'startTime', width: 18 },
      { header: '종료시간(HH:mm)', key: 'endTime', width: 18 },
      { header: '비고', key: 'memo', width: 25 },
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.height = 28;
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
      cell.font = { bold: true, color: { argb: 'FF1E293B' }, size: 11 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    const sampleRows = [
      { type: '수업', period: 1, name: '1교시', startTime: '09:00', endTime: '09:50', memo: '50분 수업 / 10분 휴식' },
      { type: '수업', period: 1, name: '1교시(선택)', startTime: '09:30', endTime: '10:20', memo: '동일 교시 시간차 예시' },
      { type: '수업', period: 2, name: '2교시', startTime: '10:00', endTime: '10:50', memo: '' },
      { type: '수업', period: 3, name: '3교시', startTime: '11:00', endTime: '11:50', memo: '' },
      { type: '수업', period: 4, name: '4교시', startTime: '12:00', endTime: '12:50', memo: '' },
      { type: '휴식/점심', period: '', name: '점심시간', startTime: '12:50', endTime: '13:40', memo: '점심시간 50분' },
      { type: '수업', period: 5, name: '5교시', startTime: '13:40', endTime: '14:30', memo: '' },
      { type: '수업', period: 6, name: '6교시', startTime: '14:40', endTime: '15:30', memo: '' },
      { type: '수업', period: 7, name: '7교시', startTime: '15:40', endTime: '16:30', memo: '' },
      { type: '수업', period: 8, name: '8교시', startTime: '16:40', endTime: '17:30', memo: '기본 8교시' },
      { type: '수업', period: 9, name: '9교시', startTime: '17:40', endTime: '18:30', memo: '야간 교시 (최대 21:50)' },
    ];

    sampleRows.forEach((row) => {
      const addedRow = worksheet.addRow(row);
      addedRow.height = 24;
      addedRow.eachCell((cell) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        this.applyBorder(cell);
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, '일과표_설정_양식.xlsx');
  },

  // 2. 일과표 엑셀 업로드 파싱
  async parsePeriodExcel(file) {
    const workbook = new ExcelJS.Workbook();
    const buffer = await file.arrayBuffer();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.worksheets[0];

    const periods = [];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const type = String(row.getCell(1).value || '').trim();
      const periodVal = row.getCell(2).value;
      const name = String(row.getCell(3).value || '').trim();
      let startTime = String(row.getCell(4).value || '').trim();
      let endTime = String(row.getCell(5).value || '').trim();

      if (!name && !startTime) return;

      const isBreak = type.includes('점심') || type.includes('휴식') || name.includes('점심') || name.includes('휴식');
      const numPeriod = isBreak ? null : (parseInt(periodVal, 10) || (periods.filter(p => !p.isBreak).length + 1));

      startTime = this.formatTimeString(startTime);
      endTime = this.formatTimeString(endTime);

      periods.push({
        id: `p_uploaded_${Date.now()}_${rowNumber}`,
        period: numPeriod,
        name: name || (isBreak ? '점심시간' : `${numPeriod}교시`),
        startTime: startTime || '09:00',
        endTime: endTime || '09:50',
        isBreak: isBreak
      });
    });

    if (periods.length === 0) {
      throw new Error('유효한 일과표 데이터를 찾을 수 없습니다.');
    }
    return periods;
  },

  // 3. 수강생 명단 및 시간표 통합 표준 양식 다운로드 (일원화 양식)
  // 3. 수강생 명단 및 시간표 통합 표준 양식 다운로드 (1·2학기 전체 포함 빈 양식)
  async downloadStudentTemplate(academicYear = '2026') {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = '강원온라인학교 시간표 시스템';
    const worksheet = workbook.addWorksheet(`${academicYear}학년도_1·2학기_통합데이터`, { views: [{ showGridLines: true }] });

    worksheet.columns = [
      { header: '순번', key: 'index', width: 8 },
      { header: '학년도', key: 'year', width: 10 },
      { header: '학기', key: 'term', width: 10 },
      { header: '학년', key: 'grade', width: 8 },
      { header: '반', key: 'classNum', width: 8 },
      { header: '번호', key: 'studentNum', width: 8 },
      { header: '학번', key: 'studentCode', width: 12 },
      { header: '성명', key: 'name', width: 14 },
      { header: '과목명', key: 'subject', width: 22 },
      { header: '개설상태', key: 'status', width: 12 },
      { header: '배정요일', key: 'day', width: 12 },
      { header: '배정교시', key: 'periodName', width: 14 },
      { header: '시작시간', key: 'startTime', width: 12 },
      { header: '종료시간', key: 'endTime', width: 12 },
      { header: '시간중복', key: 'conflict', width: 22 },
      { header: '비고', key: 'memo', width: 16 },
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
      cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    // 1학기 빈 서식 10줄 (순번 1~10)
    for (let i = 1; i <= 10; i++) {
      const addedRow = worksheet.addRow({
        index: i,
        year: academicYear,
        term: '1학기',
        grade: '',
        classNum: '',
        studentNum: '',
        studentCode: '',
        name: '',
        subject: '',
        status: '',
        day: '',
        periodName: '',
        startTime: '',
        endTime: '',
        conflict: '',
        memo: ''
      });
      addedRow.height = 24;
      addedRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { name: '맑은 고딕', size: 10 };
        this.applyBorder(cell);
        if (colNum === 3) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
          cell.font = { bold: true, color: { argb: 'FF0369A1' } };
        }
      });
    }

    // 2학기 빈 서식 10줄 (순번 11~20)
    for (let i = 11; i <= 20; i++) {
      const addedRow = worksheet.addRow({
        index: i,
        year: academicYear,
        term: '2학기',
        grade: '',
        classNum: '',
        studentNum: '',
        studentCode: '',
        name: '',
        subject: '',
        status: '',
        day: '',
        periodName: '',
        startTime: '',
        endTime: '',
        conflict: '',
        memo: ''
      });
      addedRow.height = 24;
      addedRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { name: '맑은 고딕', size: 10 };
        this.applyBorder(cell);
        if (colNum === 3) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDE9FE' } };
          cell.font = { bold: true, color: { argb: 'FF6D28D9' } };
        }
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, `${academicYear}학년도_1·2학기_시간표_통합데이터_표준양식.xlsx`);
  },

  // 4. 수강생 명단 및 시간표 통합 엑셀 업로드 파싱 (통합 양식 & 나이스 양식 완벽 호환)
  async parseStudentExcel(file) {
    const workbook = new ExcelJS.Workbook();
    const buffer = await file.arrayBuffer();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.worksheets[0];

    let headerRowIndex = -1;
    let colIndex = -1;
    let colYear = -1;
    let colSemester = -1;
    let colGrade = -1;
    let colClass = -1;
    let colNumber = -1;
    let colStudentCode = -1;
    let colName = -1;
    const colSubjects = [];
    let colStatus = -1;
    let colDay = -1;
    let colPeriod = -1;
    let colStartTime = -1;
    let colEndTime = -1;
    let colMemo = -1;

    // 헤더 행 탐색 (최대 15행 스캔)
    const maxScanRows = Math.min(15, worksheet.rowCount);
    for (let r = 1; r <= maxScanRows; r++) {
      const row = worksheet.getRow(r);
      const cellEntries = [];
      row.eachCell((cell, colNum) => {
        const text = String(cell.value || '').trim();
        if (text) cellEntries.push({ colNum, text });
      });

      const hasGrade = cellEntries.some(c => c.text === '학년' || c.text.toLowerCase() === 'grade');
      const hasCode = cellEntries.some(c => c.text === '학번' || c.text === '학생번호' || c.text.toLowerCase() === 'studentcode' || c.text.toLowerCase() === 'id');
      const hasName = cellEntries.some(c => c.text === '성명' || c.text === '이름' || c.text === '학생명' || c.text.toLowerCase() === 'name');

      if ((hasGrade || hasCode) && hasName) {
        headerRowIndex = r;
        cellEntries.forEach(c => {
          const t = c.text;
          const lower = t.toLowerCase();
          if (t === '순번' || lower === 'no' || lower === 'idx' || lower === 'index') {
            colIndex = c.colNum;
          } else if (t === '학년도' || t === '연도' || lower === 'year') {
            colYear = c.colNum;
          } else if (t === '학기' || lower === 'term' || lower === 'semester') {
            colSemester = c.colNum;
          } else if (t === '학년' || lower === 'grade') {
            colGrade = c.colNum;
          } else if (t === '반' || lower === 'class') {
            colClass = c.colNum;
          } else if (t === '번호' || t === '번' || lower === 'studentnum' || (lower === 'no' && colIndex !== c.colNum)) {
            colNumber = c.colNum;
          } else if (t === '학번' || t === '학생번호' || lower === 'studentcode' || lower === 'id') {
            colStudentCode = c.colNum;
          } else if (t === '성명' || t === '이름' || t === '학생명' || lower === 'name') {
            colName = c.colNum;
          } else if (t.includes('과목') || t.includes('교과') || t.includes('신청') || t.includes('수강') || t.includes('희망')) {
            colSubjects.push(c.colNum);
          } else if (t === '개설상태' || t === '상태' || t === '구분' || lower === 'status') {
            colStatus = c.colNum;
          } else if (t === '배정요일' || t === '요일' || lower === 'day') {
            colDay = c.colNum;
          } else if (t === '배정교시' || t === '교시' || t === '교시명' || lower === 'period') {
            colPeriod = c.colNum;
          } else if (t === '시작시간' || t === '시작' || lower === 'starttime' || lower === 'start') {
            colStartTime = c.colNum;
          } else if (t === '종료시간' || t === '종료' || lower === 'endtime' || lower === 'end') {
            colEndTime = c.colNum;
          } else if (t.includes('중복') || lower.includes('conflict')) {
            // 시간중복 열은 안내용이므로 과목명이나 기타 항목으로 오인되지 않도록 안전하게 분기
          } else if (t === '비고' || t === '강의실' || t === '개설강의실' || lower === 'memo' || lower === 'room') {
            colMemo = c.colNum;
          }
        });
        break;
      }
    }

    // 헤더를 못 찾은 경우 기본 NEIS 수강편성 양식 위치 적용
    if (headerRowIndex === -1) {
      headerRowIndex = 2;
      colSubjects.push(5); // 개설과목(학점)
      colMemo = 6; // 개설강의실
      colGrade = 3;
      colClass = 8;
      colNumber = 9;
      colName = 10;
    }

    // 학생별 통합 Map (학번_성명 기준)
    const studentMap = new Map();

    for (let r = headerRowIndex + 1; r <= worksheet.rowCount; r++) {
      const row = worksheet.getRow(r);
      const rawName = colName > 0 ? String(row.getCell(colName).value || '').trim() : '';
      if (!rawName || rawName === '성명' || rawName === '이름' || rawName === '학생명') continue;

      let grade = 1;
      let classNum = 1;
      let studentNum = null;
      let studentCode = '';

      const rawCode = colStudentCode > 0 ? String(row.getCell(colStudentCode).value || '').trim() : '';
      const rawGrade = colGrade > 0 ? row.getCell(colGrade).value : null;
      const rawClass = colClass > 0 ? row.getCell(colClass).value : null;
      const rawNumber = colNumber > 0 ? row.getCell(colNumber).value : null;

      if (rawGrade !== null && rawGrade !== undefined && String(rawGrade).trim() !== '') {
        grade = parseInt(String(rawGrade).replace(/[^0-9]/g, ''), 10) || 1;
      }
      if (rawClass !== null && rawClass !== undefined && String(rawClass).trim() !== '') {
        classNum = parseInt(String(rawClass).replace(/[^0-9]/g, ''), 10) || 1;
      }
      if (rawNumber !== null && rawNumber !== undefined && String(rawNumber).trim() !== '') {
        studentNum = parseInt(String(rawNumber).replace(/[^0-9]/g, ''), 10) || null;
      }

      // 학번 열이 존재하고 4~5자리 숫자인 경우 학년/반/번호 보정
      const cleanRawCode = rawCode.replace(/[^0-9]/g, '');
      if (cleanRawCode.length >= 4 && cleanRawCode.length <= 5) {
        if (!rawGrade && cleanRawCode.length === 5) grade = parseInt(cleanRawCode[0], 10) || grade;
        if (!rawClass && cleanRawCode.length === 5) classNum = parseInt(cleanRawCode.slice(1, 3), 10) || classNum;
        if (!studentNum && cleanRawCode.length === 5) studentNum = parseInt(cleanRawCode.slice(3), 10) || studentNum;
        studentCode = cleanRawCode;
      } else {
        studentCode = this.getStudentCode({ grade, classNum, studentNum });
      }

      // 상태 추출
      const rawStatusVal = colStatus > 0 ? row.getCell(colStatus).value : null;
      const normStatus = this.normalizeStatus(rawStatusVal);

      // 요일 추출 ('미입력', '-', '미배정' 등은 null 처리)
      let parsedDay = null;
      if (colDay > 0) {
        const rawDayVal = String(row.getCell(colDay).value || '').trim();
        if (rawDayVal && rawDayVal !== '미입력' && rawDayVal !== '미배정' && rawDayVal !== '-' && rawDayVal !== '없음') {
          parsedDay = this.dayNameToId(rawDayVal);
        }
      }

      // 교시 추출 ('미입력', '-', '미배정' 등은 null 처리)
      let parsedPeriod = null;
      let parsedPeriodName = null;
      if (colPeriod > 0) {
        const rawPeriodVal = String(row.getCell(colPeriod).value || '').trim();
        if (rawPeriodVal && rawPeriodVal !== '미입력' && rawPeriodVal !== '미배정' && rawPeriodVal !== '-' && rawPeriodVal !== '없음') {
          const num = parseInt(rawPeriodVal.replace(/[^0-9]/g, ''), 10);
          if (!isNaN(num) && num > 0) parsedPeriod = num;
          parsedPeriodName = rawPeriodVal;
        }
      }

      // 시작/종료 시간 추출
      let parsedStartTime = null;
      let parsedEndTime = null;
      if (colStartTime > 0) {
        const rawST = String(row.getCell(colStartTime).value || '').trim();
        if (rawST && rawST !== '미입력' && rawST !== '미배정' && rawST !== '-') {
          parsedStartTime = this.formatTimeString(rawST);
        }
      }
      if (colEndTime > 0) {
        const rawET = String(row.getCell(colEndTime).value || '').trim();
        if (rawET && rawET !== '미입력' && rawET !== '미배정' && rawET !== '-') {
          parsedEndTime = this.formatTimeString(rawET);
        }
      }

      // 학년도 / 학기 추출
      let parsedYear = null;
      let parsedTerm = null;
      if (colYear > 0) {
        const rawY = String(row.getCell(colYear).value || '').replace(/[^0-9]/g, '');
        if (rawY.length === 4) parsedYear = rawY;
      }
      if (colSemester > 0) {
        const rawT = String(row.getCell(colSemester).value || '').replace(/[^0-9]/g, '');
        if (rawT === '1' || rawT === '2') parsedTerm = rawT;
      }

      // 비고 추출
      let parsedMemo = colMemo > 0 ? String(row.getCell(colMemo).value || '').trim() : '';
      if (parsedMemo === '-' || parsedMemo === '미입력') parsedMemo = '';

      const uniqueKey = `${studentCode}_${rawName}`;

      // 과목 컬럼들에서 과목명 추출
      const rowSubjects = [];
      colSubjects.forEach(colIdx => {
        const val = row.getCell(colIdx).value;
        if (!val) return;
        const text = String(val).trim();
        if (!text || text === '-' || text === '미배정' || text === '없음' || text === '미입력') return;

        // 쉼표, 슬래시, 플러스, 개행 등으로 복수 과목 구분
        const tokens = text.split(/[,/+\n\r]/);
        tokens.forEach(token => {
          let cleanSub = token.trim();
          if (!cleanSub) return;
          // 뒤쪽의 학점 표기 (예: (3), (3학점), [3학점]) 제거
          cleanSub = cleanSub.replace(/[\(\[]\s*\d+\s*(?:학점)?\s*[\)\]]$/, '').trim();
          if (cleanSub && cleanSub !== '-' && cleanSub !== '미입력' && !rowSubjects.includes(cleanSub)) {
            rowSubjects.push(cleanSub);
          }
        });
      });

      if (rowSubjects.length > 0) {
        rowSubjects.forEach(cleanSub => {
          const scheduleEntry = {
            subject: cleanSub,
            status: normStatus.scheduleStatus,
            courseStatus: normStatus.courseStatus,
            statusLabel: normStatus.label,
            day: parsedDay,
            period: parsedPeriod,
            periodName: parsedPeriodName,
            startTime: parsedStartTime,
            endTime: parsedEndTime,
            year: parsedYear,
            semester: parsedTerm,
            memo: parsedMemo
          };

          if (studentMap.has(uniqueKey)) {
            const existing = studentMap.get(uniqueKey);
            if (!existing.subjects.includes(cleanSub)) {
              existing.subjects.push(cleanSub);
            }
            if (!existing.schedules) existing.schedules = [];
            existing.schedules.push(scheduleEntry);
          } else {
            studentMap.set(uniqueKey, {
              id: `s_${Date.now()}_${r}_${Math.random().toString(36).substr(2, 4)}`,
              grade: grade,
              classNum: classNum,
              studentNum: studentNum,
              studentCode: studentCode,
              name: rawName,
              subject: cleanSub,
              subjects: [cleanSub],
              schedules: [scheduleEntry],
              day: parsedDay,
              period: parsedPeriod,
              periodId: null,
              status: normStatus.scheduleStatus,
              year: parsedYear,
              semester: parsedTerm,
              memo: parsedMemo
            });
          }
        });
      } else {
        if (!studentMap.has(uniqueKey)) {
          studentMap.set(uniqueKey, {
            id: `s_${Date.now()}_${r}_${Math.random().toString(36).substr(2, 4)}`,
            grade: grade,
            classNum: classNum,
            studentNum: studentNum,
            studentCode: studentCode,
            name: rawName,
            subject: '',
            subjects: [],
            schedules: [],
            day: null,
            period: null,
            periodId: null,
            status: 'confirmed',
            year: parsedYear,
            semester: parsedTerm,
            memo: parsedMemo
          });
        }
      }
    }

    const students = Array.from(studentMap.values());
    if (students.length === 0) {
      throw new Error('유효한 학생 데이터를 찾을 수 없습니다. [학년, 반, 번호, 성명] 또는 [학번, 성명, 과목] 열을 확인해주세요.');
    }
    return students;
  },

  // 5. 결괏값 형태 1: 학급 시간표 형태 (파스텔톤 디자인 멀티시트 엑셀)
  async exportTimetableExcel(periods, days, students, subjectColorMap, academicYear = '2026', semester = '1', schoolName = '강원온라인학교') {
    const workbook = new ExcelJS.Workbook();
    const yearLabel = `${academicYear}학년도 ${semester}학기`;
    const filePrefix = `${academicYear}학년도_${semester}학기_${schoolName}`;
    workbook.creator = `${schoolName} 시간표 시스템`;

    // 5-1. 종합 시간표 시트 생성
    const totalSheet = workbook.addWorksheet('종합 시간표', { views: [{ showGridLines: true }] });
    this.buildTimetableSheet(totalSheet, periods, days, students, subjectColorMap, `${yearLabel} ${schoolName} 전체 종합 시간표`.trim());

    // 5-2. 학년-반별 개별 시트 자동 생성
    const classSet = new Set();
    students.forEach(s => {
      if (s.grade && s.classNum) {
        classSet.add(`${s.grade}-${s.classNum}`);
      }
    });

    const sortedClasses = Array.from(classSet).sort((a, b) => {
      const [gA, cA] = a.split('-').map(Number);
      const [gB, cB] = b.split('-').map(Number);
      return gA === gB ? cA - cB : gA - gB;
    });

    for (const cls of sortedClasses) {
      const [grade, classNum] = cls.split('-').map(Number);
      const sheetName = `${grade}학년 ${classNum}반`;
      const classStudents = students.filter(s => s.grade === grade && s.classNum === classNum);
      const classSheet = workbook.addWorksheet(sheetName, { views: [{ showGridLines: true }] });
      this.buildTimetableSheet(
        classSheet,
        periods,
        days,
        classStudents,
        subjectColorMap,
        `${yearLabel} ${schoolName} [${grade}학년 ${classNum}반] 시간표`.trim()
      );
    }

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, `${filePrefix}_학급_시간표_${this.getTimestamp()}.xlsx`);
  },

  // 개별 시간표 시트 디자인 생성 헬퍼
  buildTimetableSheet(sheet, periods, days, studentList, subjectColorMap, titleText) {
    sheet.columns = [
      { width: 22 },
      ...days.map(() => ({ width: 28 }))
    ];

    // 타이틀 행
    sheet.mergeCells(1, 1, 1, days.length + 1);
    const titleCell = sheet.getCell(1, 1);
    titleCell.value = titleText;
    titleCell.font = { name: '맑은 고딕', size: 15, bold: true, color: { argb: 'FF1E293B' } };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.getRow(1).height = 40;

    // 헤더 행 (요일)
    const headerRow = sheet.getRow(2);
    headerRow.height = 30;
    const firstHeader = headerRow.getCell(1);
    firstHeader.value = '교시 / 시간';
    firstHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
    firstHeader.font = { name: '맑은 고딕', bold: true, color: { argb: 'FF334155' } };
    firstHeader.alignment = { vertical: 'middle', horizontal: 'center' };
    this.applyBorder(firstHeader);

    days.forEach((day, idx) => {
      const cell = headerRow.getCell(idx + 2);
      cell.value = day.isWeekend ? `${day.name} (주말)` : `${day.name}요일`;
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: day.isWeekend ? 'FFFEE2E2' : 'FFDBEAFE' }
      };
      cell.font = {
        name: '맑은 고딕',
        bold: true,
        color: { argb: day.isWeekend ? 'FF991B1B' : 'FF1E40AF' }
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    let currentExcelRow = 3;

    periods.forEach((period) => {
      const row = sheet.getRow(currentExcelRow);
      row.height = period.isBreak ? 28 : 55;

      const periodCell = row.getCell(1);
      periodCell.value = period.isBreak
        ? `[${period.name}]\n${period.startTime}~${period.endTime}`
        : `${period.name}\n(${period.startTime}~${period.endTime})`;
      periodCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: period.isBreak ? 'FFF1F5F9' : 'FFF8FAFC' }
      };
      periodCell.font = {
        name: '맑은 고딕',
        bold: true,
        size: period.isBreak ? 9 : 10,
        color: { argb: period.isBreak ? 'FF64748B' : 'FF334155' }
      };
      periodCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      this.applyBorder(periodCell);

      if (period.isBreak) {
        sheet.mergeCells(currentExcelRow, 2, currentExcelRow, days.length + 1);
        const breakCell = row.getCell(2);
        breakCell.value = `☕ ${period.name} (${period.startTime} ~ ${period.endTime})`;
        breakCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        breakCell.font = { name: '맑은 고딕', size: 10, color: { argb: 'FF94A3B8' }, italic: true };
        breakCell.alignment = { vertical: 'middle', horizontal: 'center' };
        for (let c = 2; c <= days.length + 1; c++) {
          this.applyBorder(row.getCell(c));
        }
      } else {
        days.forEach((day, dayIdx) => {
          const cell = row.getCell(dayIdx + 2);
          
          const assigned = studentList.filter(s => {
            if (s.day !== day.id) return false;
            if (s.periodId) return s.periodId === period.id;
            return s.period === period.period;
          });

          if (assigned.length > 0) {
            // 셀 배경은 흰색으로 통일하여 특정 과목 하나의 색으로 전체가 덮이지 않도록 함
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFFFFFF' }
            };

            // 셀 내 각 과목별 텍스트 음영(폰트 컬러) 처리 (ExcelJS RichText 활용)
            const richTextItems = [];

            assigned.forEach((s, aIdx) => {
              if (aIdx > 0) {
                richTextItems.push({ text: '\n' });
              }

              const label = this.getTimetableCardLabel(s);
              const isCanceled = s.status === 'canceled' || s.status === '폐강';

              if (isCanceled) {
                // 폐강 과목: 은은한 회색 + 취소선
                richTextItems.push({
                  text: label,
                  font: {
                    name: '맑은 고딕',
                    size: assigned.length > 2 ? 8.5 : 9.5,
                    strike: true,
                    color: { argb: 'FF94A3B8' }
                  }
                });
              } else {
                // 개설/조정 과목: 과목 고유의 선명한 글자색 + 굵게 (흰 배경에서 완벽한 가독성)
                const subName = s.subject || '';
                const colorInfo = subjectColorMap[subName] || { excelBg: 'E0F2FE', excelText: '0369A1' };
                const textColor = colorInfo.excelText || '0369A1';

                richTextItems.push({
                  text: label,
                  font: {
                    name: '맑은 고딕',
                    bold: true,
                    size: assigned.length > 2 ? 8.5 : 9.5,
                    color: { argb: 'FF' + textColor }
                  }
                });
              }
            });

            cell.value = { richText: richTextItems };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

            const neededHeight = Math.max(55, assigned.length * 22);
            if (row.height < neededHeight) {
              row.height = neededHeight;
            }
          } else {
            cell.value = '';
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
          }
          this.applyBorder(cell);
        });
      }

      currentExcelRow++;
    });
  },

  // -------------------------------------------------------------
  // 시간 중복 판별 헬퍼 (동일 요일/교시에 2개 이상의 유효 과목 수강 판정)
  // 반환: Map<number, string[]> (행 인덱스 -> 충돌 상대 과목명 배열)
  // -------------------------------------------------------------
  buildConflictMap(students) {
    const conflictMap = new Map();
    const slotMap = new Map(); // key -> Array<{ subject, idx }>

    (students || []).forEach((s, idx) => {
      const isAssigned = s.day && (s.periodId || s.period);
      if (!isAssigned) return;
      const normStatus = this.normalizeStatus(s.status);
      if (normStatus.label === '폐강') return; // 폐강 과목은 제외

      const studentKey = s.studentCode || this.getStudentCode(s) || s.studentMasterId || `${s.grade}_${s.classNum}_${s.studentNum}_${s.name}`;
      const periodKey = s.periodId || s.period;
      const slotKey = `${studentKey}__${s.day}__${periodKey}`;

      if (!slotMap.has(slotKey)) {
        slotMap.set(slotKey, []);
      }
      slotMap.get(slotKey).push({ subject: s.subject, idx });
    });

    slotMap.forEach((entries) => {
      if (entries.length > 1) {
        entries.forEach(curr => {
          const others = entries.filter(e => e.subject !== curr.subject).map(e => e.subject);
          const uniqueOthers = Array.from(new Set(others.filter(Boolean)));
          if (uniqueOthers.length > 0) {
            conflictMap.set(curr.idx, uniqueOthers);
          }
        });
      }
    });

    return conflictMap;
  },

  // 6. 결괏값 형태 2: 엑셀데이터 형태 (원천 데이터 테이블 / 통합 데이터 내보내기)
  async exportDataExcel(students, periods, days, academicYear = '2026', semester = '1', schoolName = '강원온라인학교') {
    const workbook = new ExcelJS.Workbook();
    const yearLabel = `${academicYear}학년도 ${semester}학기`;
    const filePrefix = `${academicYear}학년도_${semester}학기_${schoolName}`;
    workbook.creator = `${schoolName} 시간표 시스템`;
    const worksheet = workbook.addWorksheet(`${academicYear}학년도_${semester}학기_통합데이터`, { views: [{ showGridLines: true }] });

    worksheet.columns = [
      { header: '순번', key: 'index', width: 8 },
      { header: '학년도', key: 'year', width: 10 },
      { header: '학기', key: 'term', width: 8 },
      { header: '학년', key: 'grade', width: 8 },
      { header: '반', key: 'classNum', width: 8 },
      { header: '번호', key: 'studentNum', width: 8 },
      { header: '학번', key: 'studentCode', width: 12 },
      { header: '성명', key: 'name', width: 14 },
      { header: '과목명', key: 'subject', width: 22 },
      { header: '개설상태', key: 'status', width: 12 },
      { header: '배정요일', key: 'day', width: 12 },
      { header: '배정교시', key: 'periodName', width: 14 },
      { header: '시작시간', key: 'startTime', width: 12 },
      { header: '종료시간', key: 'endTime', width: 12 },
      { header: '시간중복', key: 'conflict', width: 22 },
      { header: '비고', key: 'memo', width: 16 },
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
      cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    const dayNameMap = { mon: '월요일', tue: '화요일', wed: '수요일', thu: '목요일', fri: '금요일', sat: '토요일', sun: '일요일' };
    const conflictMap = this.buildConflictMap(students);

    students.forEach((s, idx) => {
      const isAssigned = s.day && (s.periodId || s.period);
      const periodObj = periods.find(p => p.id === s.periodId || (!s.periodId && p.period === s.period));
      const periodName = isAssigned ? (periodObj ? periodObj.name : (s.period ? `${s.period}교시` : '미입력')) : '미입력';
      const startTime = isAssigned ? (periodObj ? periodObj.startTime : '미입력') : '미입력';
      const endTime = isAssigned ? (periodObj ? periodObj.endTime : '미입력') : '미입력';
      const dayText = isAssigned ? (dayNameMap[s.day] || s.day) : '미입력';

      const studentCode = this.getStudentCode(s);
      const normStatus = this.normalizeStatus(s.status);
      const statusLabel = normStatus.label;

      const conflicting = conflictMap.get(idx) || [];
      const conflictText = conflicting.length > 0 ? `⚠️ 중복: ${conflicting.join(', ')}` : '';

      const row = worksheet.addRow({
        index: idx + 1,
        year: academicYear,
        term: `${semester}학기`,
        grade: s.grade ? `${s.grade}` : '1',
        classNum: s.classNum ? `${s.classNum}` : '1',
        studentNum: s.studentNum ? `${s.studentNum}` : '',
        studentCode: studentCode,
        name: s.name,
        subject: s.subject || '미입력',
        status: statusLabel,
        day: dayText,
        periodName: periodName,
        startTime: startTime,
        endTime: endTime,
        conflict: conflictText,
        memo: s.memo || ''
      });

      row.height = 24;
      row.eachCell((cell, colNum) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { name: '맑은 고딕', size: 10 };
        this.applyBorder(cell);

        // 개설상태 열 (col 10)
        if (colNum === 10) {
          if (statusLabel === '확정') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
            cell.font = { bold: true, color: { argb: 'FF065F46' } };
          } else if (statusLabel === '조정') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
            cell.font = { bold: true, color: { argb: 'FFB45309' } };
          } else if (statusLabel === '폐강') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
            cell.font = { strike: true, color: { argb: 'FF64748B' } };
          }
        }

        // 미입력 셀 하이라이트 (배정요일, 배정교시, 시작시간, 종료시간)
        if ([11, 12, 13, 14].includes(colNum) && cell.value === '미입력') {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
          cell.font = { color: { argb: 'FF94A3B8' } };
        }

        // 시간중복 열 (col 15)
        if (colNum === 15 && conflictText) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE4E6' } };
          cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFE11D48' }, size: 9 };
        }
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, `${filePrefix}_시간표_통합데이터_${this.getTimestamp()}.xlsx`);
  },

  // 7. 학년도 전체 (1학기 + 2학기) 통합 엑셀 다운로드 (작성 데이터 포함)
  async exportYearFullDataExcel(academicYear = '2026', semesterList = [], schoolName = '강원온라인학교') {
    const workbook = new ExcelJS.Workbook();
    const filePrefix = `${academicYear}학년도_1·2학기_${schoolName}`;
    workbook.creator = `${schoolName} 시간표 시스템`;
    const worksheet = workbook.addWorksheet(`${academicYear}학년도_전체(1·2학기)_통합데이터`, { views: [{ showGridLines: true }] });

    worksheet.columns = [
      { header: '순번', key: 'index', width: 8 },
      { header: '학년도', key: 'year', width: 10 },
      { header: '학기', key: 'term', width: 10 },
      { header: '학년', key: 'grade', width: 8 },
      { header: '반', key: 'classNum', width: 8 },
      { header: '번호', key: 'studentNum', width: 8 },
      { header: '학번', key: 'studentCode', width: 12 },
      { header: '성명', key: 'name', width: 14 },
      { header: '과목명', key: 'subject', width: 22 },
      { header: '개설상태', key: 'status', width: 12 },
      { header: '배정요일', key: 'day', width: 12 },
      { header: '배정교시', key: 'periodName', width: 14 },
      { header: '시작시간', key: 'startTime', width: 12 },
      { header: '종료시간', key: 'endTime', width: 12 },
      { header: '시간중복', key: 'conflict', width: 22 },
      { header: '비고', key: 'memo', width: 16 },
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
      cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    const dayNameMap = { mon: '월요일', tue: '화요일', wed: '수요일', thu: '목요일', fri: '금요일', sat: '토요일', sun: '일요일' };

    let globalRowIdx = 1;

    semesterList.forEach(({ semester, students, periods }) => {
      const conflictMap = this.buildConflictMap(students);

      (students || []).forEach((s, sIdx) => {
        const isAssigned = s.day && (s.periodId || s.period);
        const periodObj = (periods || []).find(p => p.id === s.periodId || (!s.periodId && p.period === s.period));
        const periodName = isAssigned ? (periodObj ? periodObj.name : (s.period ? `${s.period}교시` : '미입력')) : '미입력';
        const startTime = isAssigned ? (periodObj ? periodObj.startTime : '미입력') : '미입력';
        const endTime = isAssigned ? (periodObj ? periodObj.endTime : '미입력') : '미입력';
        const dayText = isAssigned ? (dayNameMap[s.day] || s.day) : '미입력';

        const studentCode = this.getStudentCode(s);
        const normStatus = this.normalizeStatus(s.status);
        const statusLabel = normStatus.label;

        const conflicting = conflictMap.get(sIdx) || [];
        const conflictText = conflicting.length > 0 ? `⚠️ 중복: ${conflicting.join(', ')}` : '';

        const row = worksheet.addRow({
          index: globalRowIdx++,
          year: academicYear,
          term: `${semester}학기`,
          grade: s.grade ? `${s.grade}` : '1',
          classNum: s.classNum ? `${s.classNum}` : '1',
          studentNum: s.studentNum ? `${s.studentNum}` : '',
          studentCode: studentCode,
          name: s.name,
          subject: s.subject || '미입력',
          status: statusLabel,
          day: dayText,
          periodName: periodName,
          startTime: startTime,
          endTime: endTime,
          conflict: conflictText,
          memo: s.memo || ''
        });

        row.height = 24;
        row.eachCell((cell, colNum) => {
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.font = { name: '맑은 고딕', size: 10 };
          this.applyBorder(cell);

          // 개설상태 열 (col 10)
          if (colNum === 10) {
            if (statusLabel === '확정') {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
              cell.font = { bold: true, color: { argb: 'FF065F46' } };
            } else if (statusLabel === '조정') {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
              cell.font = { bold: true, color: { argb: 'FFB45309' } };
            } else if (statusLabel === '폐강') {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
              cell.font = { strike: true, color: { argb: 'FF64748B' } };
            }
          }

          // 학기 열 구분 강조 (1학기: 연한 하늘, 2학기: 연한 보라)
          if (colNum === 3) {
            if (String(semester) === '1') {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
              cell.font = { bold: true, color: { argb: 'FF0369A1' } };
            } else {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDE9FE' } };
              cell.font = { bold: true, color: { argb: 'FF6D28D9' } };
            }
          }

          // 미입력 셀 하이라이트
          if ([11, 12, 13, 14].includes(colNum) && cell.value === '미입력') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
            cell.font = { color: { argb: 'FF94A3B8' } };
          }

          // 시간중복 열 (col 15)
          if (colNum === 15 && conflictText) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE4E6' } };
            cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFE11D48' }, size: 9 };
          }
        });
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, `${filePrefix}_시간표_통합데이터_${this.getTimestamp()}.xlsx`);
  },

  // -------------------------------------------------------------
  // 8. 1단계 전용: 학생 및 과목 편제 명단 엑셀 다운로드 (1학기 + 2학기 연속 출력 & 옅은 파스텔 음영 구분)
  // -------------------------------------------------------------
  async exportStudentRosterExcel(optionsOrStudents, maybeEnrollments, academicYear = '2026', semester = '1', schoolName = '강원온라인학교') {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = `${schoolName} 시간표 시스템`;

    let year = academicYear;
    let gradeLabel = '';
    let semestersData = [];

    // 옵션 객체 형태인지 레거시 단일 학기 호출인지 판별
    if (optionsOrStudents && typeof optionsOrStudents === 'object' && !Array.isArray(optionsOrStudents) && (optionsOrStudents.sem1Students || optionsOrStudents.sem2Students || optionsOrStudents.semesters)) {
      const opts = optionsOrStudents;
      year = opts.academicYear || year;
      gradeLabel = opts.grade ? `${opts.grade}학년` : '';
      schoolName = opts.schoolName || schoolName;

      if (Array.isArray(opts.semesters)) {
        semestersData = opts.semesters;
      } else {
        // 기본 1학기 + 2학기 순서 구조 (1학기 먼저, 2학기 이어서)
        semestersData = [
          { semester: '1', students: opts.sem1Students || [], enrollments: opts.sem1Enrollments || {} },
          { semester: '2', students: opts.sem2Students || [], enrollments: opts.sem2Enrollments || {} }
        ];
      }
    } else {
      // 레거시 단일 학기 호출 대응
      semestersData = [
        { semester: String(semester), students: optionsOrStudents || [], enrollments: maybeEnrollments || {} }
      ];
    }

    const titleGrade = gradeLabel ? `_${gradeLabel}` : '';
    const filePrefix = `${year}학년도${titleGrade}_1·2학기통합_${schoolName}`;
    const sheetName = `${year}학년도${titleGrade}_학생명단`.slice(0, 31);
    const worksheet = workbook.addWorksheet(sheetName, { views: [{ showGridLines: true }] });

    worksheet.columns = [
      { header: '순번', key: 'index', width: 8 },
      { header: '학년도', key: 'year', width: 10 },
      { header: '학기', key: 'term', width: 10 },
      { header: '학년', key: 'grade', width: 8 },
      { header: '반', key: 'classNum', width: 8 },
      { header: '번호', key: 'studentNum', width: 8 },
      { header: '학번', key: 'studentCode', width: 12 },
      { header: '성명', key: 'name', width: 14 },
      { header: '편제상태', key: 'enrollStatus', width: 12 },
      { header: '수강과목수', key: 'courseCount', width: 12 },
      { header: '수강과목목록', key: 'courses', width: 40 }
    ];

    // 헤더 행 스타일링
    const headerRow = worksheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } }; // 세련된 딥 슬레이트
      cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    let globalRowIdx = 1;

    // 학기별 옅은 파스텔톤 테마 (1학기: 연한 하늘색, 2학기: 연한 보라색)
    const semesterTheme = {
      '1': {
        // 1학기: 연한 하늘색 계열 (Sky)
        rowBg: 'FFF0F9FF',        // 행 전체 은은한 연하늘색
        termBg: 'FFE0F2FE',       // 학기 셀 연하늘 뱃지
        termText: 'FF0369A1'      // 학기 셀 진파랑 볼드
      },
      '2': {
        // 2학기: 연한 보라색 계열 (Purple)
        rowBg: 'FFFAF5FF',        // 행 전체 은은한 연보라색
        termBg: 'FFEDE9FE',       // 학기 셀 연보라 뱃지
        termText: 'FF6D28D9'      // 학기 셀 진보라 볼드
      },
      default: {
        rowBg: 'FFFFFFFF',
        termBg: 'FFF1F5F9',
        termText: 'FF334155'
      }
    };

    semestersData.forEach(({ semester, students, enrollments }) => {
      const termLabel = `${semester}학기`;
      const theme = semesterTheme[String(semester)] || semesterTheme.default;

      (students || []).forEach(st => {
        const code = st.studentCode || this.getStudentCode(st);
        const courses = [];
        Object.entries(enrollments || {}).forEach(([sub, ids]) => {
          if (ids && ids.includes(st.id)) courses.push(sub);
        });

        const isEnrolled = courses.length > 0;
        const statusLabel = isEnrolled ? '편제완료' : '미편성';
        const courseCountText = `${courses.length}과목`;
        const coursesText = isEnrolled ? courses.join(', ') : '미편성';

        const row = worksheet.addRow({
          index: globalRowIdx,
          year: year,
          term: termLabel,
          grade: st.grade ? `${st.grade}` : '1',
          classNum: st.classNum ? `${st.classNum}` : '1',
          studentNum: st.studentNum ? `${st.studentNum}` : '',
          studentCode: code,
          name: st.name || '',
          enrollStatus: statusLabel,
          courseCount: courseCountText,
          courses: coursesText
        });

        row.height = 24;
        row.eachCell((cell, colNum) => {
          cell.alignment = { vertical: 'middle', horizontal: colNum === 11 ? 'left' : 'center' };
          cell.font = { name: '맑은 고딕', size: 10, color: { argb: 'FF1E293B' } };
          this.applyBorder(cell);

          // 행 전체 옅은 파스텔 배경 적용 (1학기: 연하늘, 2학기: 연보라)
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.rowBg } };

          // 학기 열 (col 3): 학기 전용 강조 뱃지 색상
          if (colNum === 3) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.termBg } };
            cell.font = { name: '맑은 고딕', bold: true, color: { argb: theme.termText }, size: 10 };
          }

          // 편제상태 열 (col 9): 편제완료(연초록) / 미편성(연주황)
          if (colNum === 9) {
            if (isEnrolled) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FF047857' } };
            } else {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFB45309' } };
            }
          }

          // 수강과목수 열 (col 10)
          if (colNum === 10) {
            if (!isEnrolled) {
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFB45309' } };
            } else {
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: theme.termText } };
            }
          }

          // 수강과목목록 열 (col 11)
          if (colNum === 11 && !isEnrolled) {
            cell.font = { name: '맑은 고딕', color: { argb: 'FF94A3B8' }, italic: true };
          }
        });

        globalRowIdx++;
      });
    });

    // 자동 필터 적용
    worksheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: 11 }
    };

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, `${filePrefix}_학생_과목편제명단_${this.getTimestamp()}.xlsx`);
  },

  // -------------------------------------------------------------
  // 9. 학기-과목별 수강자 명단 및 인원 엑셀 다운로드 (1학기 -> 2학기 연속 출력, 학년/반/번호 정렬)
  // -------------------------------------------------------------
  async exportSubjectRosterExcel({ academicYear = '2026', schoolName = '강원온라인학교', sem1Subjects = [], sem2Subjects = [] }) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = `${schoolName} 시간표 시스템`;
    const filePrefix = `${academicYear}학년도_${schoolName}_학기_과목별_수강자명단`;

    // -----------------------------------------------------------
    // 시트 1: 과목별 상세 수강자 명단
    // -----------------------------------------------------------
    const detailSheet = workbook.addWorksheet(`${academicYear}학년도_과목별수강명단`, { views: [{ showGridLines: true }] });
    detailSheet.columns = [
      { header: '순번', key: 'index', width: 8 },
      { header: '학기', key: 'term', width: 10 },
      { header: '과목명', key: 'subject', width: 26 },
      { header: '과목 수강인원', key: 'subjectTotal', width: 14 },
      { header: '과목내 번호', key: 'subjectSeq', width: 12 },
      { header: '학년', key: 'grade', width: 8 },
      { header: '반', key: 'classNum', width: 8 },
      { header: '번호', key: 'studentNum', width: 8 },
      { header: '학번', key: 'studentCode', width: 12 },
      { header: '성명', key: 'name', width: 14 },
      { header: '상태', key: 'status', width: 12 }
    ];

    const headerRow1 = detailSheet.getRow(1);
    headerRow1.height = 30;
    headerRow1.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } }; // 세련된 딥 슬레이트
      cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    let globalRowIdx = 1;

    // 학기별 파스텔 테마 (1학기: 연한 하늘색, 2학기: 연한 보라색)
    const semesterThemes = {
      '1': {
        rowBg: 'FFF0F9FF',        // 행 전체 은은한 연하늘색
        termBg: 'FFE0F2FE',       // 학기 셀 연하늘 뱃지
        termText: 'FF0369A1',     // 학기 셀 진파랑 볼드
        subText: 'FF0C4A6E'       // 과목명 진파랑
      },
      '2': {
        rowBg: 'FFFAF5FF',        // 행 전체 은은한 연보라색
        termBg: 'FFEDE9FE',       // 학기 셀 연보라 뱃지
        termText: 'FF6D28D9',     // 학기 셀 진보라 볼드
        subText: 'FF581C87'       // 과목명 진보라
      }
    };

    const processSemesterData = (semester, subjectList) => {
      const termLabel = `${semester}학기`;
      const theme = semesterThemes[String(semester)] || semesterThemes['1'];

      (subjectList || []).forEach(subItem => {
        const subName = subItem.subject || '';
        const count = subItem.count || (subItem.students ? subItem.students.length : 0);
        const countText = `${count}명`;
        const students = subItem.students || [];
        const statusText = subItem.status === 'canceled' ? '폐강' : (subItem.status === 'adjusted' ? '조정' : '개설');

        if (students.length === 0) {
          // 수강생이 없는 과목도 명단에 1행으로 명시
          const row = detailSheet.addRow({
            index: globalRowIdx,
            term: termLabel,
            subject: subName,
            subjectTotal: '0명',
            subjectSeq: '-',
            grade: '-',
            classNum: '-',
            studentNum: '-',
            studentCode: '-',
            name: '(수강생 없음)',
            status: statusText
          });
          row.height = 24;
          row.eachCell((cell, colNum) => {
            cell.alignment = { vertical: 'middle', horizontal: colNum === 3 ? 'left' : 'center' };
            cell.font = { name: '맑은 고딕', size: 10, color: { argb: 'FF94A3B8' } };
            this.applyBorder(cell);
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.rowBg } };
            if (colNum === 2) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.termBg } };
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: theme.termText }, size: 10 };
            }
            if (colNum === 3) {
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: theme.subText }, size: 10 };
            }
          });
          globalRowIdx++;
          return;
        }

        students.forEach((st, sIdx) => {
          const code = st.studentCode || this.getStudentCode(st);
          const row = detailSheet.addRow({
            index: globalRowIdx,
            term: termLabel,
            subject: subName,
            subjectTotal: countText,
            subjectSeq: `${sIdx + 1} / ${count}`,
            grade: st.grade ? `${st.grade}` : '',
            classNum: st.classNum ? `${st.classNum}` : '',
            studentNum: st.studentNum ? `${st.studentNum}` : '',
            studentCode: code,
            name: st.name || '',
            status: statusText
          });

          row.height = 24;
          row.eachCell((cell, colNum) => {
            cell.alignment = { vertical: 'middle', horizontal: colNum === 3 ? 'left' : 'center' };
            cell.font = { name: '맑은 고딕', size: 10, color: { argb: 'FF1E293B' } };
            this.applyBorder(cell);
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.rowBg } };

            // 학기 열 (col 2)
            if (colNum === 2) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.termBg } };
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: theme.termText }, size: 10 };
            }
            // 과목명 열 (col 3)
            if (colNum === 3) {
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: theme.subText }, size: 10 };
            }
            // 과목 수강인원 열 (col 4)
            if (colNum === 4) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: theme.termBg } };
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: theme.termText }, size: 10 };
            }
            // 과목내 번호 열 (col 5)
            if (colNum === 5) {
              cell.font = { name: '맑은 고딕', size: 9, color: { argb: 'FF64748B' } };
            }
            // 성명 열 (col 10)
            if (colNum === 10) {
              cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FF0F172A' } };
            }
            // 상태 열 (col 11)
            if (colNum === 11) {
              if (subItem.status === 'canceled') {
                cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFDC2626' } };
              } else {
                cell.font = { name: '맑은 고딕', color: { argb: 'FF059669' } };
              }
            }
          });

          globalRowIdx++;
        });
      });
    };

    // 1학기 먼저 출력, 이어서 2학기 연속 출력
    processSemesterData('1', sem1Subjects);
    processSemesterData('2', sem2Subjects);

    detailSheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: 11 }
    };

    // -----------------------------------------------------------
    // 시트 2: 과목별 수강인원 한눈에 보는 요약 통계
    // -----------------------------------------------------------
    const summarySheet = workbook.addWorksheet(`${academicYear}학년도_수강인원요약`, { views: [{ showGridLines: true }] });
    summarySheet.columns = [
      { header: '순번', key: 'idx', width: 8 },
      { header: '과목명', key: 'subject', width: 28 },
      { header: '1학기 수강인원', key: 'sem1Count', width: 16 },
      { header: '2학기 수강인원', key: 'sem2Count', width: 16 },
      { header: '연간 총 수강건수', key: 'totalCount', width: 18 },
      { header: '개설 상태', key: 'status', width: 14 }
    ];

    const headerRow2 = summarySheet.getRow(1);
    headerRow2.height = 30;
    headerRow2.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } }; // 진한 에메랄드/틸
      cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    // 1학기와 2학기 과목 통합 맵 구성
    const summaryMap = new Map();
    (sem1Subjects || []).forEach(s => {
      if (!summaryMap.has(s.subject)) {
        summaryMap.set(s.subject, { subject: s.subject, sem1: s.count || 0, sem2: 0, status: s.status || 'confirmed' });
      } else {
        summaryMap.get(s.subject).sem1 = s.count || 0;
      }
    });
    (sem2Subjects || []).forEach(s => {
      if (!summaryMap.has(s.subject)) {
        summaryMap.set(s.subject, { subject: s.subject, sem1: 0, sem2: s.count || 0, status: s.status || 'confirmed' });
      } else {
        summaryMap.get(s.subject).sem2 = s.count || 0;
      }
    });

    const sortedSummary = Array.from(summaryMap.values()).sort((a, b) => a.subject.localeCompare(b.subject, 'ko'));
    let totalSem1 = 0;
    let totalSem2 = 0;

    sortedSummary.forEach((item, sIdx) => {
      const sum = item.sem1 + item.sem2;
      totalSem1 += item.sem1;
      totalSem2 += item.sem2;
      const statusLabel = item.status === 'canceled' ? '폐강' : '정상 개설';

      const row = summarySheet.addRow({
        idx: sIdx + 1,
        subject: item.subject,
        sem1Count: `${item.sem1}명`,
        sem2Count: `${item.sem2}명`,
        totalCount: `${sum}명`,
        status: statusLabel
      });

      row.height = 24;
      row.eachCell((cell, colNum) => {
        cell.alignment = { vertical: 'middle', horizontal: colNum === 2 ? 'left' : 'center' };
        cell.font = { name: '맑은 고딕', size: 10 };
        this.applyBorder(cell);

        // 줄무늬 배경
        if (sIdx % 2 === 1) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        }
        // 1학기 셀 (연하늘 폰트)
        if (colNum === 3 && item.sem1 > 0) {
          cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FF0369A1' } };
        }
        // 2학기 셀 (연보라 폰트)
        if (colNum === 4 && item.sem2 > 0) {
          cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FF6D28D9' } };
        }
        // 총 수강건수 셀
        if (colNum === 5) {
          cell.font = { name: '맑은 고딕', bold: true, color: { argb: 'FF0F172A' } };
        }
      });
    });

    // 합계 요약 행
    const sumRow = summarySheet.addRow({
      idx: '합계',
      subject: `총 ${sortedSummary.length}개 과목`,
      sem1Count: `${totalSem1}명`,
      sem2Count: `${totalSem2}명`,
      totalCount: `${totalSem1 + totalSem2}명`,
      status: '-'
    });
    sumRow.height = 28;
    sumRow.eachCell((cell, colNum) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
      cell.font = { name: '맑은 고딕', bold: true, size: 10, color: { argb: 'FF0F172A' } };
      cell.alignment = { vertical: 'middle', horizontal: colNum === 2 ? 'left' : 'center' };
      this.applyBorder(cell);
    });

    summarySheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: 6 }
    };

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, `${filePrefix}_${this.getTimestamp()}.xlsx`);
  },

  applyBorder(cell) {
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
    };
  },

  saveFile(buffer, fileName) {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    }, 100);
  },

  formatTimeString(str) {
    if (!str) return '';
    const clean = str.replace(/[^0-9:]/g, '');
    const parts = clean.split(':');
    if (parts.length >= 2) {
      const h = parts[0].padStart(2, '0');
      const m = parts[1].padStart(2, '0');
      return `${h}:${m}`;
    }
    return str;
  },

  getTimestamp() {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    const hh = String(now.getHours()).padStart(2, '0');
    const mm = String(now.getMinutes()).padStart(2, '0');
    return `${y}${m}${d}_${hh}${mm}`;
  }
};
