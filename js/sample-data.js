/**
 * 기본 일과표 및 설정 데이터
 */

const DEFAULT_PERIODS = [
  { id: 'p1', period: 1, name: '1교시', startTime: '09:00', endTime: '09:50', isBreak: false },
  { id: 'p2', period: 2, name: '2교시', startTime: '10:00', endTime: '10:50', isBreak: false },
  { id: 'p3', period: 3, name: '3교시', startTime: '11:00', endTime: '11:50', isBreak: false },
  { id: 'p4', period: 4, name: '4교시', startTime: '12:00', endTime: '12:50', isBreak: false },
  { id: 'p_lunch', period: null, name: '점심시간', startTime: '12:50', endTime: '13:40', isBreak: true },
  { id: 'p5', period: 5, name: '5교시', startTime: '13:40', endTime: '14:30', isBreak: false },
  { id: 'p6', period: 6, name: '6교시', startTime: '14:40', endTime: '15:30', isBreak: false },
  { id: 'p7', period: 7, name: '7교시', startTime: '15:40', endTime: '16:30', isBreak: false },
  { id: 'p8', period: 8, name: '8교시', startTime: '16:40', endTime: '17:30', isBreak: false },
];

const DEFAULT_DAYS = [
  { id: 'mon', name: '월', fullName: '월요일', isWeekend: false },
  { id: 'tue', name: '화', fullName: '화요일', isWeekend: false },
  { id: 'wed', name: '수', fullName: '수요일', isWeekend: false },
  { id: 'thu', name: '목', fullName: '목요일', isWeekend: false },
  { id: 'fri', name: '금', fullName: '금요일', isWeekend: false },
  { id: 'sat', name: '토', fullName: '토요일', isWeekend: true },
  { id: 'sun', name: '일', fullName: '일요일', isWeekend: true },
];

// 정선된 파스텔 색상 팔레트 풀
const PASTEL_PALETTE = [
  { bg: '#E0F2FE', text: '#0369A1', border: '#BAE6FD', excelBg: 'E0F2FE', excelText: '0369A1' }, // 하늘
  { bg: '#D1FAE5', text: '#047857', border: '#A7F3D0', excelBg: 'D1FAE5', excelText: '047857' }, // 에메랄드
  { bg: '#EDE9FE', text: '#6D28D9', border: '#DDD6FE', excelBg: 'EDE9FE', excelText: '6D28D9' }, // 보라
  { bg: '#FFE4E6', text: '#BE123C', border: '#FECDD3', excelBg: 'FFE4E6', excelText: 'BE123C' }, // 로즈
  { bg: '#FEF3C7', text: '#B45309', border: '#FDE68A', excelBg: 'FEF3C7', excelText: 'B45309' }, // 앰버
  { bg: '#FFEDD5', text: '#C2410C', border: '#FED7AA', excelBg: 'FFEDD5', excelText: 'C2410C' }, // 살구
  { bg: '#CCFBF1', text: '#0F766E', border: '#99F6E4', excelBg: 'CCFBF1', excelText: '0F766E' }, // 틸
  { bg: '#FCE7F3', text: '#BE185D', border: '#FBCFE8', excelBg: 'FCE7F3', excelText: 'BE185D' }, // 핑크
  { bg: '#E0E7FF', text: '#4338CA', border: '#C7D2FE', excelBg: 'E0E7FF', excelText: '4338CA' }, // 인디고
  { bg: '#ECFCCB', text: '#4D7C0F', border: '#D9F99D', excelBg: 'ECFCCB', excelText: '4D7C0F' }, // 라임
  { bg: '#CFFAFE', text: '#0E7490', border: '#A5F3FC', excelBg: 'CFFAFE', excelText: '0E7490' }, // 시안
  { bg: '#F3E8FF', text: '#7E22CE', border: '#E9D5FF', excelBg: 'F3E8FF', excelText: '7E22CE' }, // 퍼플
  { bg: '#FEE2E2', text: '#B91C1C', border: '#FECACA', excelBg: 'FEE2E2', excelText: 'B91C1C' }, // 레드
  { bg: '#FEF9C3', text: '#A16207', border: '#FEF08A', excelBg: 'FEF9C3', excelText: 'A16207' }, // 옐로우
  { bg: '#E2E8F0', text: '#334155', border: '#CBD5E1', excelBg: 'E2E8F0', excelText: '334155' }  // 슬레이트
];

// 기존 예시 학생 데이터 삭제 - 빈 목록으로 시작
const INITIAL_STUDENTS = [];

// 기본 개설 과목 - 빈 목록으로 시작 (사용자 정의 또는 엑셀 업로드로 구성)
const DEFAULT_OFFERED_SUBJECTS = [];
