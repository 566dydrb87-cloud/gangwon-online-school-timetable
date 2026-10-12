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

  // 3. 수강생 명단 표준 양식 다운로드
  async downloadStudentTemplate() {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = '강원온라인학교 시간표 시스템';
    const worksheet = workbook.addWorksheet('편성학생 명단');

    worksheet.columns = [
      { header: '학년도', key: 'year', width: 10 },
      { header: '학기', key: 'term', width: 8 },
      { header: '학년', key: 'grade', width: 10 },
      { header: '편제명', key: 'track', width: 14 },
      { header: '개설과목(학점)', key: 'subject', width: 22 },
      { header: '개설강의실', key: 'room', width: 14 },
      { header: '계열/학년/학과', key: 'dept', width: 18 },
      { header: '반', key: 'classNum', width: 8 },
      { header: '번호', key: 'studentNum', width: 8 },
      { header: '성명', key: 'name', width: 16 },
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.height = 28;
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } };
      cell.font = { bold: true, color: { argb: 'FF1E3A8A' }, size: 11 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      this.applyBorder(cell);
    });

    const samples = [
      { year: 2026, term: 1, grade: 1, track: '선택교육과정', subject: '인공지능 기초(3)', room: '온라인1실', dept: '일반/1/공통', classNum: 1, studentNum: 1, name: '홍길동' },
      { year: 2026, term: 1, grade: 1, track: '선택교육과정', subject: '인공지능 기초(3)', room: '온라인1실', dept: '일반/1/공통', classNum: 1, studentNum: 2, name: '이순신' },
      { year: 2026, term: 1, grade: 2, track: '선택교육과정', subject: '빅데이터 분석(3)', room: '온라인2실', dept: '일반/2/공통', classNum: 2, studentNum: 5, name: '강감찬' },
    ];

    samples.forEach((row) => {
      const addedRow = worksheet.addRow(row);
      addedRow.height = 24;
      addedRow.eachCell((cell) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        this.applyBorder(cell);
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, '수강신청_편성학생_양식.xlsx');
  },

  // 4. 수강생 명단 엑셀 업로드 파싱 (요청사항: 학번과 성명만 읽고 표시)
  async parseStudentExcel(file) {
    const workbook = new ExcelJS.Workbook();
    const buffer = await file.arrayBuffer();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.worksheets[0];

    const students = [];
    let headerRowIndex = -1;
    let colGrade = -1;
    let colClass = -1;
    let colNumber = -1;
    let colName = -1;

    // 헤더 행 탐색
    const maxScanRows = Math.min(15, worksheet.rowCount);
    for (let r = 1; r <= maxScanRows; r++) {
      const row = worksheet.getRow(r);
      const cellEntries = [];
      row.eachCell((cell, colNum) => {
        const text = String(cell.value || '').trim();
        if (text) cellEntries.push({ colNum, text });
      });

      const hasGrade = cellEntries.some(c => c.text === '학년');
      const hasName = cellEntries.some(c => c.text === '성명' || c.text === '이름' || c.text === '학생명');

      if (hasGrade && hasName) {
        headerRowIndex = r;
        cellEntries.forEach(c => {
          const t = c.text;
          if (t === '학년') colGrade = c.colNum;
          else if (t === '반') colClass = c.colNum;
          else if (t === '번호') colNumber = c.colNum;
          else if (t === '성명' || t === '이름' || t === '학생명') colName = c.colNum;
        });
        break;
      }
    }

    if (headerRowIndex === -1) {
      headerRowIndex = 2;
      colGrade = 3;
      colClass = 8;
      colNumber = 9;
      colName = 10;
    }

    // 중복 방지 Set (학번_성명 기준)
    const seen = new Set();

    for (let r = headerRowIndex + 1; r <= worksheet.rowCount; r++) {
      const row = worksheet.getRow(r);
      const rawName = colName > 0 ? String(row.getCell(colName).value || '').trim() : '';
      if (!rawName || rawName === '성명' || rawName === '이름') continue;

      const rawGrade = colGrade > 0 ? row.getCell(colGrade).value : 1;
      const rawClass = colClass > 0 ? row.getCell(colClass).value : 1;
      const rawNumber = colNumber > 0 ? row.getCell(colNumber).value : null;

      const grade = parseInt(String(rawGrade).replace(/[^0-9]/g, ''), 10) || 1;
      const classNum = parseInt(String(rawClass).replace(/[^0-9]/g, ''), 10) || 1;
      const studentNum = rawNumber !== null && rawNumber !== undefined
        ? (parseInt(String(rawNumber).replace(/[^0-9]/g, ''), 10) || null)
        : null;

      const studentCode = this.getStudentCode({ grade, classNum, studentNum });
      const uniqueKey = `${studentCode}_${rawName}`;

      // 중복 학생은 1명만 등록 (이후 교사가 여러 과목 배정 가능)
      if (seen.has(uniqueKey)) continue;
      seen.add(uniqueKey);

      // 과목은 아직 배정되지 않은 상태 ('')로 학번과 성명만 가져옴
      students.push({
        id: `s_${Date.now()}_${r}`,
        grade: grade,
        classNum: classNum,
        studentNum: studentNum,
        studentCode: studentCode,
        name: rawName,
        subject: '', // 업로드 시점에는 과목 미배정
        day: null,
        period: null,
        periodId: null,
        status: 'confirmed' // 기본 상태: 확정(🟢)
      });
    }

    if (students.length === 0) {
      throw new Error('유효한 학생 데이터를 찾을 수 없습니다. [학년, 반, 번호, 성명] 열을 확인해주세요.');
    }
    return students;
  },

  // 5. 결괏값 형태 1: 학급 시간표 형태 (파스텔톤 디자인 멀티시트 엑셀)
  async exportTimetableExcel(periods, days, students, subjectColorMap, schoolName = '강원온라인학교') {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = `${schoolName} 시간표 시스템`;

    // 5-1. 종합 시간표 시트 생성
    const totalSheet = workbook.addWorksheet('종합 시간표', { views: [{ showGridLines: true }] });
    this.buildTimetableSheet(totalSheet, periods, days, students, subjectColorMap, `${schoolName} 전체 종합 시간표`);

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
        `${schoolName} [${grade}학년 ${classNum}반] 시간표`
      );
    }

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, `${schoolName}_학급_시간표_${this.getTimestamp()}.xlsx`);
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
            // 시간표 셀 표기: 상태기호 과목명(학번성명)
            const textLines = assigned.map(s => this.getTimetableCardLabel(s));
            cell.value = textLines.join('\n');
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

            const firstItem = assigned[0];
            const isCanceled = firstItem.status === 'canceled';

            if (isCanceled) {
              // 폐강인 경우 회색 배경 + 취소선
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE2E8F0' } // 회색
              };
              cell.font = {
                name: '맑은 고딕',
                strike: true,
                size: assigned.length > 2 ? 8 : 9,
                color: { argb: 'FF64748B' }
              };
            } else {
              // 확정 및 조정: 고유 파스텔 배경
              const firstSubject = firstItem.subject;
              const colorInfo = subjectColorMap[firstSubject] || { excelBg: 'E0F2FE', excelText: '0369A1' };
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF' + colorInfo.excelBg }
              };
              cell.font = {
                name: '맑은 고딕',
                bold: true,
                size: assigned.length > 2 ? 8 : 9,
                color: { argb: 'FF' + colorInfo.excelText }
              };
            }

            const neededHeight = Math.max(55, assigned.length * 20);
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

  // 6. 결괏값 형태 2: 엑셀데이터 형태 (원천 데이터 테이블)
  async exportDataExcel(students, periods, days, schoolName = '강원온라인학교') {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = `${schoolName} 시간표 시스템`;
    const worksheet = workbook.addWorksheet('수강배정_원천데이터', { views: [{ showGridLines: true }] });

    worksheet.columns = [
      { header: '순번', key: 'index', width: 8 },
      { header: '학년', key: 'grade', width: 8 },
      { header: '반', key: 'classNum', width: 8 },
      { header: '번호', key: 'studentNum', width: 8 },
      { header: '학번', key: 'studentCode', width: 12 },
      { header: '성명', key: 'name', width: 14 },
      { header: '학번성명', key: 'codeName', width: 16 },
      { header: '과목명', key: 'subject', width: 22 },
      { header: '시간표 표기', key: 'timetableLabel', width: 26 },
      { header: '상태', key: 'statusText', width: 14 },
      { header: '배정요일', key: 'day', width: 12 },
      { header: '배정교시', key: 'periodName', width: 14 },
      { header: '수업시간', key: 'timeRange', width: 18 },
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

    students.forEach((s, idx) => {
      const isAssigned = s.day && (s.periodId || s.period);
      const periodObj = periods.find(p => p.id === s.periodId || (!s.periodId && p.period === s.period));
      const periodName = periodObj ? periodObj.name : (s.period ? `${s.period}교시` : '-');
      const timeRange = periodObj ? `${periodObj.startTime} ~ ${periodObj.endTime}` : '-';

      const studentCode = this.getStudentCode(s);
      const codeName = this.getStudentCodeName(s);
      const timetableLabel = isAssigned ? this.getTimetableCardLabel(s) : '-';
      const statusText = isAssigned ? this.getStatusText(s.status) : '미배정';

      const row = worksheet.addRow({
        index: idx + 1,
        grade: s.grade ? `${s.grade}` : '-',
        classNum: s.classNum ? `${s.classNum}` : '-',
        studentNum: s.studentNum ? `${s.studentNum}` : '-',
        studentCode: studentCode,
        name: s.name,
        codeName: codeName,
        subject: s.subject || '-',
        timetableLabel: timetableLabel,
        statusText: statusText,
        day: isAssigned ? (dayNameMap[s.day] || s.day) : '-',
        periodName: isAssigned ? periodName : '-',
        timeRange: isAssigned ? timeRange : '-'
      });

      row.height = 24;
      row.eachCell((cell, colNum) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { name: '맑은 고딕', size: 10 };
        this.applyBorder(cell);

        // 상태 열 하이라이트
        if (colNum === 10) {
          if (!isAssigned) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
            cell.font = { color: { argb: 'FF64748B' } };
          } else if (s.status === 'confirmed') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
            cell.font = { bold: true, color: { argb: 'FF065F46' } };
          } else if (s.status === 'adjust') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
            cell.font = { bold: true, color: { argb: 'FFB45309' } };
          } else if (s.status === 'canceled') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
            cell.font = { strike: true, color: { argb: 'FF64748B' } };
          }
        }
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    this.saveFile(buffer, `${schoolName}_수강데이터_원천목록_${this.getTimestamp()}.xlsx`);
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
