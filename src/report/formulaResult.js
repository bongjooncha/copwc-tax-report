// 수식 칸의 계산 결과 보관
// ExcelJS(4.4.0)는 수식 값 { formula, result }를 저장할 때 result가 0·빈 문자열·false면 버림
// (lib/doc/cell.js FormulaValue._copyModel의 truthy 검사). 그러면 미리보기·PDF에서 0 칸이 빈칸이 됨.
// 수식 값을 넣을 때 결과를 시트별로 따로 보관하고, 읽을 때 ExcelJS 값에 결과가 없으면 보관한 값을 씀.
// (xlsx 파일은 fullCalcOnLoad로 Excel이 열 때 다시 계산하므로 영향 없음)

const RESULTS = new WeakMap(); // worksheet → Map(address → result)

const isFormula = (v) => v !== null && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v);

export function setCellValue(cell, value) {
  cell.value = value;
  const ws = cell.worksheet;
  let m = RESULTS.get(ws);
  if (isFormula(value) && value.result !== undefined) {
    if (!m) RESULTS.set(ws, (m = new Map()));
    m.set(cell.address, value.result);
  } else if (m) {
    m.delete(cell.address);
  }
  return cell;
}

/** 칸의 표시 값: 수식이면 계산 결과(ExcelJS가 버린 0 포함), 아니면 값 그대로 */
export function cellResult(cell) {
  const v = cell.value;
  if (!isFormula(v)) return v;
  if (v.result !== undefined) return v.result;
  const m = RESULTS.get(cell.worksheet);
  return m && m.has(cell.address) ? m.get(cell.address) : undefined;
}
