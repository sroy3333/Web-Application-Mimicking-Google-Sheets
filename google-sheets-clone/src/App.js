import React, { useState, useEffect, useRef, useCallback } from "react";
import Handsontable from "handsontable";
import "handsontable/dist/handsontable.full.css";
import axios from "axios";
import { evaluate } from "mathjs";
import "./App.css";

const API_URL = "http://localhost:5000";

export default function Spreadsheet() {
  const hotRef = useRef(null);
  const [spreadsheetName, setSpreadsheetName] = useState("mySheet");

  const initializeTable = useCallback((data) => {
    if (hotRef.current) {
      hotRef.current.destroy();
    }

    hotRef.current = new Handsontable(document.getElementById("spreadsheet"), {
      data,
      rowHeaders: true,
      colHeaders: true,
      minCols: 10,
      minRows: 10,
      contextMenu: true,
      formulas: true,
      afterChange: (changes) => {
        if (changes) {
          handleFormulaParsing();
        }
      },
      licenseKey: "non-commercial-and-evaluation",
    });
  }, []);

  useEffect(() => {
    axios.get(`${API_URL}/load/${spreadsheetName}`).then((response) => {
      const data = response.data.length ? response.data : [[""]];
      initializeTable(data);
    });
  }, [spreadsheetName, initializeTable]);

  const handleFormulaParsing = () => {
    const hot = hotRef.current;
    if (!hot) return;

    const data = hot.getData();

    data.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        if (typeof cell === "string" && cell.startsWith("=")) {
          try {
            let formula = cell.substring(1);

            const extractValues = (startCell, endCell) => {
              const startCol = startCell.charCodeAt(0) - 65; // Convert column letter to index
              const startRow = parseInt(startCell.substring(1), 10) - 1;
              const endCol = endCell.charCodeAt(0) - 65;
              const endRow = parseInt(endCell.substring(1), 10) - 1;
            
              let values = [];
              for (let i = startRow; i <= endRow; i++) {
                for (let j = startCol; j <= endCol; j++) {
                  let cellValue = data[i]?.[j];
            
                  // ✅ Check if the value is a valid number before adding
                  if (cellValue !== undefined && cellValue !== null && cellValue !== '' && !isNaN(cellValue)) {
                    values.push(Number(cellValue));
                  }
                }
              }
              return values;
            };
            
            

            formula = formula.replace(/SUM\((\w+\d+):(\w+\d+)\)/g, (_, startCell, endCell) => {
              return extractValues(startCell, endCell).reduce((acc, val) => acc + val, 0);
            });

            formula = formula.replace(/AVERAGE\((\w+\d+):(\w+\d+)\)/g, (_, startCell, endCell) => {
              let values = extractValues(startCell, endCell);
              return values.length > 0 ? values.reduce((a, b) => a + b, 0) / values.length : 0;
            });

            formula = formula.replace(/MAX\((\w+\d+):(\w+\d+)\)/g, (_, startCell, endCell) => {
              return Math.max(...extractValues(startCell, endCell));
            });

            formula = formula.replace(/MIN\((\w+\d+):(\w+\d+)\)/g, (_, startCell, endCell) => {
              return Math.min(...extractValues(startCell, endCell));
            });

            formula = formula.replace(/COUNT\((\w+\d+):(\w+\d+)\)/g, (_, startCell, endCell) => {
              return extractValues(startCell, endCell).filter((val) => typeof val === "number" && !isNaN(val)).length;
            });
            

            formula = formula.replace(/TRIM\((\w+\d+)\)/g, (_, cell) => {
              const colIndex = cell.charCodeAt(0) - 65;
              const rowIndex = parseInt(cell.substring(1), 10) - 1;
              return (data[rowIndex]?.[colIndex] || "").trim();
            });

            formula = formula.replace(/UPPER\((\w+\d+)\)/g, (_, cell) => {
              const colIndex = cell.charCodeAt(0) - 65;
              const rowIndex = parseInt(cell.substring(1), 10) - 1;
              let cellValue = data[rowIndex]?.[colIndex];
              return typeof cellValue === "string" ? cellValue.toUpperCase() : "ERROR";
            });

            formula = formula.replace(/LOWER\((\w+\d+)\)/g, (_, cell) => {
              const colIndex = cell.charCodeAt(0) - 65;
              const rowIndex = parseInt(cell.substring(1), 10) - 1;
              return (data[rowIndex]?.[colIndex] || "").toString().toLowerCase();
            });

            formula = formula.replace(/FIND_AND_REPLACE\((\w+\d+),\s*"(.*?)",\s*"(.*?)"\)/g, (_, cell, findText, replaceText) => {
              const colIndex = cell.charCodeAt(0) - 65; // Convert column letter to index (A=0, B=1, etc.)
              const rowIndex = parseInt(cell.substring(1), 10) - 1; // Convert row number (1-based) to index (0-based)
              
              let cellValue = data[rowIndex]?.[colIndex] || "";
            
              let newValue = cellValue.toString().replace(new RegExp(findText, "g"), replaceText);
            
              // Update only the referenced cell (A1 in this case)
              hot.setDataAtCell(rowIndex, colIndex, newValue);
            
              // Prevent formula cell (A2) from displaying the output
              return "";
            });
            
            

            formula = formula.replace(/REMOVE_DUPLICATES\((\w+\d+):(\w+\d+)\)/g, (_, startCell, endCell) => {
              const startCol = startCell.charCodeAt(0) - 65;
              const startRow = parseInt(startCell.substring(1), 10) - 1;
              const endCol = endCell.charCodeAt(0) - 65;
              const endRow = parseInt(endCell.substring(1), 10) - 1;

              let seen = new Set();
              for (let i = startRow; i <= endRow; i++) {
                let cellValue = data[i]?.[startCol];
                if (seen.has(cellValue)) {
                  hot.setDataAtCell(i, startCol, ""); // Remove duplicates
                } else {
                  seen.add(cellValue);
                }
              }
              return "";
            });

            
            

            formula = formula.replace(/([A-Z]+)(\d+)/g, (_, col, row) => {
              const colIndex = col.charCodeAt(0) - 65;
              const rowIndex = parseInt(row, 10) - 1;
              let cellValue = data[rowIndex]?.[colIndex];
              
              if (typeof cellValue === "string" && !cellValue.startsWith('"') && !cellValue.startsWith("=")) {
                return `"${cellValue}"`;  // Ensure text values are enclosed in double quotes
              }
              
              return cellValue ?? 0;
            });

            let result = evaluate(formula);
            hot.setDataAtCell(rowIndex, colIndex, result);
          } catch (error) {
            hot.setDataAtCell(rowIndex, colIndex, "ERROR");
          }
        }
      });
    });
  };

  const saveSpreadsheet = () => {
    const data = hotRef.current.getData();
    axios.post(`${API_URL}/save`, { name: spreadsheetName, data }).then(() => {
      alert("Spreadsheet saved!");
    });
  };

  return (
    <div className="p-4">
      <h1 className="text-2xl font-bold">Google Sheets Clone</h1>
      <input type="text" value={spreadsheetName} onChange={(e) => setSpreadsheetName(e.target.value)} className="border p-2 mr-2" placeholder="Spreadsheet Name" />
      <button onClick={saveSpreadsheet} className="bg-blue-500 text-white px-4 py-2 rounded">Save</button>
      <div id="spreadsheet" className="mt-4"></div>
    </div>
  );
}