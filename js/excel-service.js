/**
 * Excel Service using ExcelJS for rich styling, pastel colors, and multi-sheet support.
 */

const ExcelService = {
  // 요일 텍스트를 day id로 변환하는 유틸
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

  // 시간표 표기 형태: 과목명(학번이름) -> 인공지능(10101홍길동)
  getTimetableCardLabel(student) {
    const sub = student.subject || '과목';
    const codeName = this.getStudentCodeName(student);
    return `${sub}(${codeName})`;
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
      if (rowNumber === 1) return; // 헤더 스킵
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
      throw new Error('유효한 일과표 데이터를 찾을 수 없습니다. 양식을 확인해주세요.');
    }
    return periods;
  },

  // 3. 수강생 명단 표준 양식 다운로드 (업로드 파일 포맷 안내)
  async downloadStudentTemplate() {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = '강원온라인학교 시간표 시스템';
    const worksheet = workbook.addWorksheet('편성학생 명단');

    // 사용자가 제공한 NEIS / 온라인학교 표준 양식 컬럼 적용
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

  // 4. 수강생 명단 엑셀 업로드 파싱 (사용자 NEIS 형식 완벽 호환)
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
    let colSubject = -1;

    // 헤더 행 자동 감지 (1행~15행 스캔)
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
          else if (t.includes('개설과목') || t === '과목명' || t === '과목') colSubject = c.colNum;
        });
        break;
      }
    }

    // 헤더를 못 찾은 경우 기본 컬럼 위치 (사용자 명시 레이아웃 기준)
    // 1:학년도, 2:학기, 3:학년, 4:편제명, 5:개설과목(학점), 6:개설강의실, 7:계열..., 8:반, 9:번호, 10:성명
    if (headerRowIndex === -1) {
      headerRowIndex = 2;
      colGrade = 3;
      colSubject = 5;
      colClass = 8;
      colNumber = 9;
      colName = 10;
    }

    // 학생 데이터 파싱
    for (let r = headerRowIndex + 1; r <= worksheet.rowCount; r++) {
      const row = worksheet.getRow(r);
      const rawName = colName > 0 ? String(row.getCell(colName).value || '').trim() : '';
      if (!rawName || rawName === '성명' || rawName === '이름') continue;

      const rawGrade = colGrade > 0 ? row.getCell(colGrade).value : 1;
      const rawClass = colClass > 0 ? row.getCell(colClass).value : 1;
      const rawNumber = colNumber > 0 ? row.getCell(colNumber).value : null;
      let rawSubject = colSubject > 0 ? String(row.getCell(colSubject).value || '').trim() : '';

      // 과목명 정제: '인공지능 기초(3)' -> '인공지능 기초', '(3학점)' 등 단위 제거
      if (rawSubject) {
        rawSubject = rawSubject.replace(/\s*\(\d+학점\)/g, '').replace(/\s*\(\d+\)/g, '').trim();
      }

      const grade = parseInt(String(rawGrade).replace(/[^0-9]/g, ''), 10) || 1;
      const classNum = parseInt(String(rawClass).replace(/[^0-9]/g, ''), 10) || 1;
      const studentNum = rawNumber !== null && rawNumber !== undefined
        ? (parseInt(String(rawNumber).replace(/[^0-9]/g, ''), 10) || null)
        : null;

      // 배정 상태는 알 필요 없이 초기 미배정 상태로 생성
      students.push({
        id: `s_${Date.now()}_${r}`,
        grade: grade,
        classNum: classNum,
        studentNum: studentNum,
        name: rawName,
        subject: rawSubject || '미지정',
        day: null,
        period: null,
        periodId: null
      });
    }

    if (students.length === 0) {
      throw new Error('유효한 학생 데이터를 찾을 수 없습니다. [학년, 반, 번호, 성명] 열이 포함되어 있는지 확인해주세요.');
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
      { width: 22 }, // 교시/시간 열
      ...days.map(() => ({ width: 26 })) // 요일 열들
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
          
          // periodId 우선 매칭, 없을 경우 period 번호 매칭
          const assigned = studentList.filter(s => {
            if (s.day !== day.id) return false;
            if (s.periodId) return s.periodId === period.id;
            return s.period === period.period;
          });

          if (assigned.length > 0) {
            // 시간표 셀 표기: 과목명(학번이름) -> 인공지능(10101홍길동)
            const textLines = assigned.map(s => this.getTimetableCardLabel(s));
            cell.value = textLines.join('\n');
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

            const firstSubject = assigned[0].subject;
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
      { header: '시간표 표기', key: 'timetableLabel', width: 24 },
      { header: '배정상태', key: 'status', width: 12 },
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
      const timetableLabel = this.getTimetableCardLabel(s);

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
        status: isAssigned ? '배정완료' : '미배정',
        day: isAssigned ? (dayNameMap[s.day] || s.day) : '-',
        periodName: isAssigned ? periodName : '-',
        timeRange: isAssigned ? timeRange : '-'
      });

      row.height = 24;
      row.eachCell((cell, colNum) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { name: '맑은 고딕', size: 10 };
        this.applyBorder(cell);

        if (colNum === 10) { // 배정상태 하이라이트
          if (isAssigned) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
            cell.font = { bold: true, color: { argb: 'FF065F46' } };
          } else {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
            cell.font = { bold: true, color: { argb: 'FF991B1B' } };
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
