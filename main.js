class RangeSnapshot {
  /**
   * @param {number} startRow - 1-indexed
   * @param {number} startCol - 1-indexed
   * @param {number} numRows
   * @param {number} numCols
   * @param {any[][]} values - getValues()で取得した2D配列
   */
  constructor(startRow, startCol, numRows, numCols, values) {
    this._startRow = startRow;
    this._startCol = startCol;
    this._numRows = numRows;
    this._numCols = numCols;
    // 内部データはディープコピーして保持
    this._values = values.map((row) => [...row]);
  }

  /**
   * A1記法をranges内の行・列インデックスに変換
   * @param {string} a1 - 例: "B2"
   * @returns {{ ri: number, ci: number } | null} - 範囲外ならnull
   */
  _a1ToIndex(a1) {
    const match = a1.toUpperCase().match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;

    const col = match[1]
      .split("")
      .reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0);
    const row = parseInt(match[2], 10);

    const ri = row - this._startRow;
    const ci = col - this._startCol;

    if (ri < 0 || this._numRows <= ri || ci < 0 || this._numCols <= ci)
      return null;

    return { ri, ci };
  }

  /**
   * 指定セルの値を取得。範囲外なら空文字を返す
   * @param {string} a1
   * @returns {any}
   */
  getValue(a1) {
    const idx = this._a1ToIndex(a1);
    if (!idx) return "";
    return this._values[idx.ri][idx.ci];
  }

  /**
   * 指定セルに値をセットした新しいRangeSnapshotを返す（イミュータブル）
   * @param {string} a1
   * @param {any} value
   * @returns {RangeSnapshot}
   */
  setValue(a1, value) {
    const idx = this._a1ToIndex(a1);
    if (!idx) return this;

    const newValues = this._values.map((row) => [...row]);
    newValues[idx.ri][idx.ci] = value;
    return new RangeSnapshot(
      this._startRow,
      this._startCol,
      this._numRows,
      this._numCols,
      newValues,
    );
  }

  /**
   * 現在保持している2D配列を返す
   * @returns {any[][]}
   */
  getValues() {
    return this._values.map((row) => [...row]);
  }

  /**
   * 保持している値をシートに書き戻す
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
   */
  apply(sheet) {
    sheet
      .getRange(this._startRow, this._startCol, this._numRows, this._numCols)
      .setValues(this._values);
  }
}
