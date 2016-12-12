import {encodeCell} from './encode-cell.js';
import cellToObject from './cell-to-object';

/**
 * Returns the XLSX-Worksheet object given the data of a
 * table calculated by `tableToData`.
 *
 * @param {object} data - The data calculated by `tableToData`.
 * @param {array} typeHandlers - The registered cell type handlers.
 *
 * @returns {object} - XLSX-Worksheet object.
 */
export default function dataToWorksheet(data, typeHandlers) {
  const { cells, ranges } = data;
  let lastColumn = 0;

  // convert cells array to an object by iterating over all rows
  const worksheet = cells.reduce((sheet, row, rowIndex) => {

    // iterate over all row cells
    row.forEach((cell, columnIndex) => {
      lastColumn = Math.max(lastColumn, columnIndex);

      // convert the row and column indices to a XLSX index
      const ref = encodeCell({
        c: columnIndex,
        r: rowIndex,
      });

      // only save actual cells and convert them to XLSX-Cell objects
      if (cell) {
        sheet[ref] = cellToObject(cell, typeHandlers);
      } else {
        sheet[ref] =  { t: 's', v: '' };
      }
    });

    return sheet;
  }, {});

  /**
   * Дублируем стили границ объекта на другой объект
   * @param target - объект, к которому нужно скопировать стили
   * @param donor - объект, от которого копируем стили
   * @param position - позиция границы
   */
  function setStyle(target, donor, position){
    "use strict";

    if (donor.s.border && position && donor.s.border[position]){
      target.s = target.s || {};
      target.s.border = target.s.border || {};
      target.s.border[position] = donor.s.border[position];
    }
  }

  ranges.forEach(range => {
    const ref = encodeCell({
      c: range.s.c,
      r: range.s.r,
    });

    let startCellObject = worksheet[ref];

    if (startCellObject.s){

      for(let sr = range.s.r; sr <= range.e.r; sr++){
        for(let sc = range.s.c; sc <= range.e.c; sc++){
          let targetCellRef = encodeCell({
            c: sc,
            r: sr,
          });
          let targetCellObject = worksheet[targetCellRef];

          if (sr == range.s.r){
            setStyle(targetCellObject, startCellObject, 'top');
          }

          if (sr == range.e.r){
            setStyle(targetCellObject, startCellObject, 'bottom');
          }

          if (sc == range.s.c){
            setStyle(targetCellObject, startCellObject, 'left');
          }

          if (sc == range.e.c){
            setStyle(targetCellObject, startCellObject, 'right');
          }
        }

      }

    }
  });

  worksheet['!cols'] = [];

  cells[0].forEach((cell, columnIndex) => {
    // only save actual cells and convert them to XLSX-Cell objects
    if (cell) {
      worksheet['!cols'].push({
        wpx: cell.offsetWidth
      });

    } else {
      worksheet['!cols'].push(null);
    }
  });

  // calculate last table index (bottom right)
  const lastRef = encodeCell({
    c: lastColumn,
    r: cells.length - 1,
  });

  // add last table index and ranges to the worksheet
  worksheet['!ref'] = `A1:${lastRef}`;
  worksheet['!merges'] = ranges;

  return worksheet;
}
