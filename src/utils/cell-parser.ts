/**
 * 单元格地址解析工具
 * 支持多字母列（如 AA、AZ）
 */

/**
 * 将列字母转换为列号（1-based）
 * A -> 1, Z -> 26, AA -> 27, AZ -> 52
 */
export function columnLetterToNumber(letters: string): number {
  let result = 0
  for (let i = 0; i < letters.length; i++) {
    result = result * 26 + (letters.charCodeAt(i) - 64)
  }
  return result
}

/**
 * 将列号转换为列字母（1-based）
 * 1 -> A, 26 -> Z, 27 -> AA
 */
export function columnNumberToLetter(num: number): string {
  let result = ''
  while (num > 0) {
    const remainder = (num - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    num = Math.floor((num - 1) / 26)
  }
  return result
}

/**
 * 解析单元格地址，返回列号和行号（均为 1-based）
 */
export function parseCellAddress(cell: string): { col: number; row: number } {
  const match = cell.match(/^([A-Z]+)(\d+)$/)
  if (!match) {
    return { col: 1, row: 1 }
  }
  return {
    col: columnLetterToNumber(match[1]),
    row: parseInt(match[2])
  }
}
