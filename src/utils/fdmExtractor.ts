import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import type { FDMItemDefinition, ExtractionResult } from '@/types/fdm';

// Item definitions - same as Python script
export const itemsDefinitions: FDMItemDefinition[] = [
  // --- BAGIAN AWAL ---
  { label: "KPP", sheet: "Sheet Home", addr: "H5", mode: "Static" },
  { label: "Sektor", sheet: "Sheet Home", addr: "D8", mode: "Static" },

  // --- BAGIAN 1: Data Umum ---
  { label: "NAMA WAJIB PAJAK", sheet: "Sheet Home", addr: "H15", mode: "Static" },
  { label: "NOMOR OBJEK PAJAK", sheet: "Sheet Home", addr: "H21", mode: "Static" },
  { label: "KELURAHAN", sheet: "Sheet Home", addr: "H27", mode: "Static" },
  { label: "KECAMATAN", sheet: "Sheet Home", addr: "H29", mode: "Static" },
  { label: "KABUPATEN/KOTA", sheet: "Sheet Home", addr: "H31", mode: "Static" },
  { label: "PROVINSI", sheet: "Sheet Home", addr: "H33", mode: "Static" },

  // Rumus Penjumlahan Areal
  { label: "LUAS BUMI", sheet: "Sheet Home", mode: "Formula_LuasBumi" },

  // --- BAGIAN 2: Areal ---
  { label: "Areal Produktif", sheet: "Sheet Home", addr: "J75", mode: "Static" },
  { label: "Areal Belum Diolah", sheet: "Sheet Home", addr: "J77", mode: "Static" },
  { label: "Areal Sudah Diolah Belum Ditanami", sheet: "Sheet Home", addr: "J78", mode: "Static" },
  { label: "Areal Pembibitan", sheet: "Sheet Home", addr: "J79", mode: "Static" },
  { label: "Areal Tidak Produktif", sheet: "Sheet Home", addr: "J80", mode: "Static" },
  { label: "Areal Pengaman", sheet: "Sheet Home", addr: "J81", mode: "Static" },
  { label: "Areal Emplasemen", sheet: "Sheet Home", addr: "J82", mode: "Static" },

  { label: "Areal Produktif (Copy)", sheet: "Sheet Home", mode: "Formula_CopyProduktif" },

  // --- BAGIAN 3: NJOP & Rumus ---
  { label: "NJOP/M Areal Belum Produktif", sheet: "C.1", addr: "BK22", mode: "Static" },

  { label: "NJOP Bumi Berupa Tanah (Rp)", sheet: "Sheet Home", mode: "Formula_NJOPTanah" },

  { label: "NJOP Bumi Berupa Pengembangan Tanah (Rp)", sheet: "C.2", keyword: "Pengembangan Tanah", mode: "Dynamic_Col_G" },

  { label: "NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)", sheet: "N/A", mode: "Formula_BIT" },

  { label: "NJOP Bumi Areal Produktif (Rp)", sheet: "N/A", mode: "Formula_NJOP_Total" },

  { label: "Luas Bumi Areal Produktif (m²)", sheet: "N/A", mode: "Formula_Luas_Ref" },

  { label: "NJOP Bumi Per M2 Areal Produktif (Rp/m2)", sheet: "N/A", mode: "Formula_NJOP_PerM2" },

  // --- BAGIAN 4: FDM Kebun ABC ---
  { label: "NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Final_Calc" },
  { label: "NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi" },

  { label: "NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI", sheet: "FDM Kebun ABC", addr: "E20", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_BelumProd" },

  // --- Tambahan Copy & NJOP Lainnya ---
  { label: "Areal Tidak Produktif (Copy)", sheet: "N/A", mode: "Formula_CopyTidakProduktif" },
  { label: "NJOP/M Areal Tidak Produktif", sheet: "C.1", addr: "BK62", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Calc_TidakProd" },
  { label: "NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_TidakProd" },

  { label: "Areal Pengaman (Copy)", sheet: "N/A", mode: "Formula_CopyPengaman" },
  { label: "NJOP/M Areal Pengaman", sheet: "D", addr: "L23", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Calc_Pengaman" },
  { label: "NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_Pengaman" },

  { label: "Areal Emplasemen (Copy)", sheet: "N/A", mode: "Formula_CopyEmplasemen" },
  { label: "NJOP/M Areal Emplasemen", sheet: "C.1", addr: "BK102", mode: "Static" },
  { label: "NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Calc_Emplasemen" },
  { label: "NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)", sheet: "N/A", mode: "Formula_Proyeksi_Emplasemen" },

  { label: "JUMLAH Luas (m2) pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Total_Luas_Ref" },
  { label: "JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI", sheet: "N/A", mode: "Formula_Total_NJOP_Sum" },

  { label: "NJOP BUMI (Rp) NJOP Bumi Per Meter Persegi pada A. DATA BUMI", sheet: "FDM Kebun ABC", addr: "E25", mode: "Static" },

  // --- BAGIAN 5: FDM Bangunan ---
  { label: "Jumlah LUAS pada B. DATA BANGUNAN", sheet: "FDM Kebun ABC", mode: "Dynamic_FDM_Bangunan_Luas" },
  { label: "Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN", sheet: "N/A", mode: "Formula_Calc_Bangunan" },
  { label: "NJOP BANGUNAN PER METER PERSEGI*) pada B. DATA BANGUNAN", sheet: "FDM Kebun ABC", mode: "Dynamic_FDM_Bangunan_PerM2" },

  { label: "TOTAL NJOP (TANAH + BANGUNAN) 2025", sheet: "N/A", mode: "Formula_Grand_Total" },

  // --- KOLOM TAMBAHAN SPPT & SIMULASI ---
  { label: "SPPT 2025", sheet: "N/A", mode: "Formula_SPPT_2025" },
  { label: "SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)", sheet: "N/A", mode: "Formula_Simulasi_NJOP_2026" },
  { label: "SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)", sheet: "N/A", mode: "Formula_Simulasi_SPPT_2026" },
  { label: "Kenaikan", sheet: "N/A", mode: "Formula_Kenaikan" },

  { label: "Persentase", sheet: "N/A", mode: "Formula_Persentase" },
  { label: "SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)", sheet: "N/A", mode: "Formula_Simulasi_Total_2026_NDT46" },
  { label: "SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)", sheet: "N/A", mode: "Formula_Simulasi_SPPT_2026_NDT46" },
];

// Helper function to get column letter from index (1-based)
function getColumnLetter(colIndex: number): string {
  let result = '';
  while (colIndex > 0) {
    colIndex--;
    result = String.fromCharCode(65 + (colIndex % 26)) + result;
    colIndex = Math.floor(colIndex / 26);
  }
  return result;
}

// Smart sheet finder - same as Python
function getSheetSmart(workbook: XLSX.WorkBook, nameHint: string): XLSX.WorkSheet | null {
  if (nameHint === "N/A") return null;
  const nameHintLower = nameHint.toLowerCase();
  const sheetMap: { [key: string]: string } = {};
  
  for (const name of workbook.SheetNames) {
    sheetMap[name.toLowerCase()] = name;
  }
  
  if (workbook.Sheets[nameHint]) return workbook.Sheets[nameHint];
  if (sheetMap[nameHintLower]) return workbook.Sheets[sheetMap[nameHintLower]];
  
  for (const existingSheet of Object.keys(sheetMap)) {
    if (nameHintLower.includes("c.1") && existingSheet.includes("c.1")) 
      return workbook.Sheets[sheetMap[existingSheet]];
    if (nameHintLower.includes("c.2") && existingSheet.includes("c.2")) 
      return workbook.Sheets[sheetMap[existingSheet]];
    if (nameHintLower.includes("home") && existingSheet.includes("home")) 
      return workbook.Sheets[sheetMap[existingSheet]];
    if (nameHintLower.includes("fdm") && existingSheet.includes("fdm")) 
      return workbook.Sheets[sheetMap[existingSheet]];
    if ((nameHintLower === "d" || nameHintLower === "sheet d") && 
        (existingSheet === "d" || existingSheet === "sheet d")) 
      return workbook.Sheets[sheetMap[existingSheet]];
  }
  return null;
}

// Get cell value from sheet
function getCellValue(sheet: XLSX.WorkSheet, addr: string): any {
  const cell = sheet[addr];
  return cell ? cell.v : null;
}

// Find FDM anchor row for dynamic extraction
function findFDMAnchorRow(sheet: XLSX.WorkSheet): number | null {
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
  
  for (let row = 19; row <= Math.min(149, range.e.r); row++) {
    for (let col = 0; col <= 4; col++) {
      const cellAddr = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = sheet[cellAddr];
      if (cell && cell.v && typeof cell.v === 'string') {
        if (cell.v.toUpperCase().includes("NJOP BANGUNAN PER METER PERSEGI")) {
          return row + 1; // Convert to 1-based
        }
      }
    }
  }
  return null;
}

// Extract data from a single file
export async function extractFDMData(file: File): Promise<ExtractionResult> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const result: ExtractionResult = {
          rowNumber: 0,
          fileName: file.name,
          data: {}
        };

        // Pre-scan FDM sheet for anchor row
        const fdmSheet = getSheetSmart(workbook, "FDM Kebun ABC");
        let fdmAnchorRow: number | null = null;
        if (fdmSheet) {
          fdmAnchorRow = findFDMAnchorRow(fdmSheet);
        }

        // Extract data for each item
        for (const item of itemsDefinitions) {
          const mode = item.mode;
          let val: any = null;

          if (mode.includes("Formula")) {
            result.data[item.label] = null; // Will be filled later with formulas
            continue;
          }

          const ws = getSheetSmart(workbook, item.sheet);

          if (!ws) {
            val = "Sheet Not Found";
          } else {
            if (mode === "Static" && item.addr) {
              try {
                val = getCellValue(ws, item.addr);
              } catch {
                val = "Error";
              }
            } else if (mode === "Dynamic_Col_G" && item.keyword) {
              val = "TIDAK DITEMUKAN";
              const keyword = item.keyword.toLowerCase();
              const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
              
              for (let row = 0; row <= Math.min(149, range.e.r); row++) {
                let found = false;
                for (let col = 0; col <= 4; col++) {
                  const cellAddr = XLSX.utils.encode_cell({ r: row, c: col });
                  const cell = ws[cellAddr];
                  if (cell && cell.v && typeof cell.v === 'string') {
                    const cellText = cell.v.toLowerCase().replace(/\s+/g, ' ');
                    if (cellText.includes(keyword)) {
                      const colGAddr = XLSX.utils.encode_cell({ r: row, c: 6 });
                      val = ws[colGAddr]?.v;
                      found = true;
                      break;
                    }
                  }
                }
                if (found) break;
              }
            } else if (mode.includes("Dynamic_FDM_Bangunan")) {
              if (fdmAnchorRow && fdmSheet) {
                if (mode === "Dynamic_FDM_Bangunan_PerM2") {
                  const addr = XLSX.utils.encode_cell({ r: fdmAnchorRow - 1, c: 4 });
                  val = fdmSheet[addr]?.v;
                } else if (mode === "Dynamic_FDM_Bangunan_NJOP") {
                  const addr = XLSX.utils.encode_cell({ r: fdmAnchorRow - 2, c: 4 });
                  val = fdmSheet[addr]?.v;
                } else if (mode === "Dynamic_FDM_Bangunan_Luas") {
                  const addr = XLSX.utils.encode_cell({ r: fdmAnchorRow - 2, c: 3 });
                  val = fdmSheet[addr]?.v;
                }
              } else {
                val = "Anchor Not Found";
              }
            }
          }

          // Clean KELURAHAN value
          if (item.label === "KELURAHAN" && val && typeof val === 'string') {
            if (val.includes("#")) {
              val = val.replace(/#\s*\d+.*$/, '').trim();
            }
          }

          result.data[item.label] = val;
        }

        resolve(result);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
}

// Generate Excel with formulas using ExcelJS for better formatting support
export async function generateResultExcel(results: ExtractionResult[]): Promise<ExcelJS.Workbook> {
  const headers = ["NO", ...itemsDefinitions.map(item => item.label)];
  
  // Create workbook
  const workbook = new ExcelJS.Workbook();
  
  // Create Sheet 1: Hasil
  const ws = workbook.addWorksheet("1. Hasil");
  
  // Set column widths
  ws.columns = headers.map(() => ({ width: 25 }));
  
  // Add headers
  const headerRow = ws.addRow(headers);
  headerRow.font = { bold: true };
  
  // Add data rows
  for (let idx = 0; idx < results.length; idx++) {
    const result = results[idx];
    const rowData: any[] = [idx + 1];
    
    for (const item of itemsDefinitions) {
      if (item.mode.includes("Formula")) {
        rowData.push(null); // Placeholder for formula
      } else {
        rowData.push(result.data[item.label] ?? null);
      }
    }
    
    ws.addRow(rowData);
  }
  
  // Add formulas for each row
  const colMap: { [key: string]: string } = {};
  for (let i = 0; i < headers.length; i++) {
    colMap[headers[i]] = getColumnLetter(i + 1);
  }

  for (let idx = 0; idx < results.length; idx++) {
    const excelRow = idx + 2;
    
    // Formula: LUAS BUMI
    const arealCols = ["Areal Produktif", "Areal Belum Diolah", "Areal Sudah Diolah Belum Ditanami", "Areal Pembibitan", "Areal Tidak Produktif", "Areal Pengaman", "Areal Emplasemen"];
    const cellsToSum = arealCols.map(c => `${colMap[c]}${excelRow}`).join(",");
    ws.getCell(`${colMap["LUAS BUMI"]}${excelRow}`).value = { formula: `SUM(${cellsToSum})` };

    // Formula: Areal Produktif (Copy)
    ws.getCell(`${colMap["Areal Produktif (Copy)"]}${excelRow}`).value = { formula: `${colMap["Areal Produktif"]}${excelRow}` };

    // Formula: NJOP Bumi Berupa Tanah (Rp)
    ws.getCell(`${colMap["NJOP Bumi Berupa Tanah (Rp)"]}${excelRow}`).value = { 
      formula: `${colMap["Areal Produktif"]}${excelRow}*${colMap["NJOP/M Areal Belum Produktif"]}${excelRow}` 
    };

    // Formula: NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)
    const colAsalBIT = colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp)"];
    ws.getCell(`${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)"]}${excelRow}`).value = { 
      formula: `${colAsalBIT}${excelRow}+(${colAsalBIT}${excelRow}*'2. Kesimpulan'!$E$2)` 
    };

    // Formula: NJOP Bumi Areal Produktif (Rp)
    const colTanah = colMap["NJOP Bumi Berupa Tanah (Rp)"];
    const colPengembangan = colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp)"];
    ws.getCell(`${colMap["NJOP Bumi Areal Produktif (Rp)"]}${excelRow}`).value = { 
      formula: `${colTanah}${excelRow}+${colPengembangan}${excelRow}` 
    };

    // Formula: Luas Bumi Areal Produktif (m²)
    ws.getCell(`${colMap["Luas Bumi Areal Produktif (m²)"]}${excelRow}`).value = { 
      formula: `${colMap["Areal Produktif"]}${excelRow}` 
    };

    // Formula: NJOP Bumi Per M2 Areal Produktif (Rp/m2)
    const colNJOPTotalProd = colMap["NJOP Bumi Areal Produktif (Rp)"];
    const colLuasProd = colMap["Luas Bumi Areal Produktif (m²)"];
    ws.getCell(`${colMap["NJOP Bumi Per M2 Areal Produktif (Rp/m2)"]}${excelRow}`).value = { 
      formula: `${colNJOPTotalProd}${excelRow}/${colLuasProd}${excelRow}` 
    };

    // Formula: NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI
    const colHargaM2 = colMap["NJOP Bumi Per M2 Areal Produktif (Rp/m2)"];
    ws.getCell(`${colMap["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI"]}${excelRow}`).value = { 
      formula: `${colLuasProd}${excelRow}*${colHargaM2}${excelRow}` 
    };

    // Formula: NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    const colBITNaik = colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)"];
    const colArealProd = colMap["Areal Produktif"];
    ws.getCell(`${colMap["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`).value = { 
      formula: `ROUND(((${colTanah}${excelRow}+${colBITNaik}${excelRow})/${colArealProd}${excelRow}),0)*${colLuasProd}${excelRow}` 
    };

    // Formula: NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    const colBelumProd = colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI"];
    ws.getCell(`${colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`).value = { 
      formula: `${colBelumProd}${excelRow}*(1+'2. Kesimpulan'!$E$14)` 
    };

    // Formula: Areal Tidak Produktif (Copy)
    const colTidakProdAsal = colMap["Areal Tidak Produktif"];
    ws.getCell(`${colMap["Areal Tidak Produktif (Copy)"]}${excelRow}`).value = { 
      formula: `${colTidakProdAsal}${excelRow}` 
    };

    // Formula: NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI
    const colTidakProdCopy = colMap["Areal Tidak Produktif (Copy)"];
    const colNJOPMTidakProd = colMap["NJOP/M Areal Tidak Produktif"];
    ws.getCell(`${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI"]}${excelRow}`).value = { 
      formula: `${colTidakProdCopy}${excelRow}*${colNJOPMTidakProd}${excelRow}` 
    };

    // Formula: NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    const colNJOPTidakProd = colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI"];
    ws.getCell(`${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`).value = { 
      formula: `${colNJOPTidakProd}${excelRow}*(1+'2. Kesimpulan'!$E$14)` 
    };

    // Formula: Areal Pengaman (Copy)
    const colPengamanAsal = colMap["Areal Pengaman"];
    ws.getCell(`${colMap["Areal Pengaman (Copy)"]}${excelRow}`).value = { 
      formula: `${colPengamanAsal}${excelRow}` 
    };

    // Formula: NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI
    const colPengamanCopy = colMap["Areal Pengaman (Copy)"];
    const colNJOPMPengaman = colMap["NJOP/M Areal Pengaman"];
    ws.getCell(`${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI"]}${excelRow}`).value = { 
      formula: `${colPengamanCopy}${excelRow}*${colNJOPMPengaman}${excelRow}` 
    };

    // Formula: NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    const colNJOPPengaman = colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI"];
    ws.getCell(`${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`).value = { 
      formula: `${colNJOPPengaman}${excelRow}*(1+'2. Kesimpulan'!$E$14)` 
    };

    // Formula: Areal Emplasemen (Copy)
    const colEmplasemenAsal = colMap["Areal Emplasemen"];
    ws.getCell(`${colMap["Areal Emplasemen (Copy)"]}${excelRow}`).value = { 
      formula: `${colEmplasemenAsal}${excelRow}` 
    };

    // Formula: NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI
    const colEmplasemenCopy = colMap["Areal Emplasemen (Copy)"];
    const colNJOPMEmplasemen = colMap["NJOP/M Areal Emplasemen"];
    ws.getCell(`${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"]}${excelRow}`).value = { 
      formula: `${colEmplasemenCopy}${excelRow}*${colNJOPMEmplasemen}${excelRow}` 
    };

    // Formula: NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)
    const colNJOPEmplasemen = colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"];
    ws.getCell(`${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`).value = { 
      formula: `${colNJOPEmplasemen}${excelRow}*(1+'2. Kesimpulan'!$E$14)` 
    };

    // Formula: JUMLAH Luas (m2) pada A. DATA BUMI
    const colLuasBumiGlobal = colMap["LUAS BUMI"];
    ws.getCell(`${colMap["JUMLAH Luas (m2) pada A. DATA BUMI"]}${excelRow}`).value = { 
      formula: `${colLuasBumiGlobal}${excelRow}` 
    };

    // Formula: JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI
    const njopComponents = ["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI", "NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"];
    const colsToSum = njopComponents.map(c => `${colMap[c]}${excelRow}`).join("+");
    ws.getCell(`${colMap["JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI"]}${excelRow}`).value = { 
      formula: colsToSum 
    };

    // Formula: Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN
    const colLuasBgn = colMap["Jumlah LUAS pada B. DATA BANGUNAN"];
    const colNJOPM2Bgn = colMap["NJOP BANGUNAN PER METER PERSEGI*) pada B. DATA BANGUNAN"];
    ws.getCell(`${colMap["Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN"]}${excelRow}`).value = { 
      formula: `${colLuasBgn}${excelRow}*${colNJOPM2Bgn}${excelRow}` 
    };

    // Formula: TOTAL NJOP (TANAH + BANGUNAN) 2025
    const colTotalBumi = colMap["JUMLAH NJOP BUMI (Rp) pada A. DATA BUMI"];
    const colTotalBgn = colMap["Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN"];
    ws.getCell(`${colMap["TOTAL NJOP (TANAH + BANGUNAN) 2025"]}${excelRow}`).value = { 
      formula: `${colTotalBumi}${excelRow}+${colTotalBgn}${excelRow}` 
    };

    // Formula: SPPT 2025
    const colTotalNJOP25 = colMap["TOTAL NJOP (TANAH + BANGUNAN) 2025"];
    ws.getCell(`${colMap["SPPT 2025"]}${excelRow}`).value = { 
      formula: `((${colTotalNJOP25}${excelRow}-12000000)*40%)*0.5%` 
    };

    // Formula: SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)
    const T = `${colMap["NJOP Bumi Berupa Tanah (Rp)"]}${excelRow}`;
    const V = `${colMap["NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT 10.3%)"]}${excelRow}`;
    const R = `${colMap["Areal Produktif"]}${excelRow}`;
    const X = `${colMap["Luas Bumi Areal Produktif (m²)"]}${excelRow}`;
    const AB = `${colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI"]}${excelRow}`;
    const AF = `${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI"]}${excelRow}`;
    const AJ = `${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI"]}${excelRow}`;
    const AN = `${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI"]}${excelRow}`;
    const AT = `${colMap["Jumlah NJOP BANGUNAN pada B. DATA BANGUNAN"]}${excelRow}`;
    ws.getCell(`${colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"]}${excelRow}`).value = { 
      formula: `(ROUND(((${T}+${V})/${R}),0)*${X})+${AB}+${AF}+${AJ}+${AN}+${AT}` 
    };

    // Formula: SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)
    const colSimNJOP26 = colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"];
    ws.getCell(`${colMap["SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"]}${excelRow}`).value = { 
      formula: `((${colSimNJOP26}${excelRow}-12000000)*40%)*0.5%` 
    };

    // Formula: Kenaikan
    const colSimSPPT26 = colMap["SIMULASI SPPT 2026 (Hanya Kenaikan BIT 10,3% + NDT Tetap)"];
    const colSPPT25 = colMap["SPPT 2025"];
    ws.getCell(`${colMap["Kenaikan"]}${excelRow}`).value = { 
      formula: `${colSimSPPT26}${excelRow}-${colSPPT25}${excelRow}` 
    };

    // Formula: Persentase
    const colKenaikan = colMap["Kenaikan"];
    ws.getCell(`${colMap["Persentase"]}${excelRow}`).value = { 
      formula: `${colKenaikan}${excelRow}/${colSPPT25}${excelRow}` 
    };

    // Formula: SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)
    const AA = `${colMap["NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AC = `${colMap["NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AG = `${colMap["NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AK = `${colMap["NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    const AO = `${colMap["NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik 46%)"]}${excelRow}`;
    ws.getCell(`${colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)"]}${excelRow}`).value = { 
      formula: `(${AA}+${AC}+${AG}+${AK}+${AO})+${AT}` 
    };

    // Formula: SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)
    const colSimNJOP26NDT = colMap["SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT 10,3% + NDT 46%)"];
    ws.getCell(`${colMap["SIMULASI SPPT 2026 (Kenaikan BIT 10,3% + NDT 46%)"]}${excelRow}`).value = { 
      formula: `((${colSimNJOP26NDT}${excelRow}-12000000)*40%)*0.5%` 
    };
  }

  // Apply number format #,##0 to columns J to BC (rows 2 onwards)
  for (let col = 10; col <= 55; col++) {
    const colLetter = getColumnLetter(col);
    for (let row = 2; row <= results.length + 1; row++) {
      const cell = ws.getCell(`${colLetter}${row}`);
      if (cell.value !== null && cell.value !== undefined) {
        cell.numFmt = '#,##0';
      }
    }
  }

  // HEADER DINAMIS SHEET 1
  const headerUpdates: { [key: string]: string } = {
    'V1': '="NJOP Bumi Berupa Pengembangan Tanah (Rp) (Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"%)"',
    'AA1': '="NJOP BUMI (Rp) AREA PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"',
    'AC1': '="NJOP BUMI (Rp) AREAL BELUM PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"',
    'AG1': '="NJOP BUMI (Rp) AREAL TIDAK PRODUKTIF pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"',
    'AK1': '="NJOP BUMI (Rp) AREAL PENGAMAN pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"',
    'AO1': '="NJOP BUMI (Rp) AREAL EMPLASEMEN pada A. DATA BUMI (Proyeksi NDT Naik "&\'2. Kesimpulan\'!$E$14*100&"%)"',
    'AX1': '="SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Hanya Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% dan NDT Tetap)"',
    'AY1': '="SIMULASI SPPT 2026 (Hanya Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% dan NDT Tetap)"',
    'BB1': '="SIMULASI TOTAL NJOP (TANAH + BANGUNAN) 2026 (Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% + NDT "&\'2. Kesimpulan\'!$E$14*100&"%)"',
    'BC1': '="SIMULASI SPPT 2026 (Kenaikan BIT "&\'2. Kesimpulan\'!$E$2*100&"% + NDT "&\'2. Kesimpulan\'!$E$14*100&"%)"'
  };

  for (const [cellAddr, formula] of Object.entries(headerUpdates)) {
    ws.getCell(cellAddr).value = { formula: formula };
  }

    // Create Sheet 2: Kesimpulan
  const ws2 = workbook.addWorksheet("2. Kesimpulan");
  
  // Set column widths
  ws2.columns = [
    { width: 60 }, { width: 30 }, { width: 25 }, { width: 20 }, { width: 20 }
  ];

  // Row 1
  ws2.getCell('E1').value = "Skenario Kenaikan BIT";
  ws2.getCell('A1').value = "Poin";
  ws2.getCell('B1').value = { formula: '"Keterangan (BIT + "&E2*100&"% dan NDT Tetap)"' };
  ws2.getCell('C1').value = "Nilai";
  ws2.getCell('D1').value = "Keterangan";
  
  // E2 - Format Persen 10.3%
  ws2.getCell('E2').value = 0.103;
  ws2.getCell('E2').numFmt = '0.0%';
  
  // Rows 2-6
  ws2.getCell('A2').value = "Simulasi Penerimaan PBB 2026";
  ws2.getCell('B2').value = "Perkebunan";
  ws2.getCell('C2').value = { formula: "SUMIF('1. Hasil'!C2:C10000,\"Perkebunan\",'1. Hasil'!AY2:AY10000)" };
  ws2.getCell('C2').numFmt = '#,##0';
  
  ws2.getCell('A3').value = "Simulasi Penerimaan PBB 2026";
  ws2.getCell('B3').value = "Minerba";
  ws2.getCell('C3').value = { formula: "SUMIF('1. Hasil'!C2:C10000,\"Minerba\",'1. Hasil'!AY2:AY10000)" };
  ws2.getCell('C3').numFmt = '#,##0';
  
  ws2.getCell('A4').value = "Simulasi Penerimaan PBB 2026";
  ws2.getCell('B4').value = "Perhutanan (HTI)";
  ws2.getCell('C4').value = { formula: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (HTI)\",'1. Hasil'!AY2:AY10000)" };
  ws2.getCell('C4').numFmt = '#,##0';
  
  ws2.getCell('A5').value = "Simulasi Penerimaan PBB 2026";
  ws2.getCell('B5').value = "Perhutanan (Hutan Alam)";
  ws2.getCell('C5').value = { formula: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (Hutan Alam)\",'1. Hasil'!AY2:AY10000)" };
  ws2.getCell('C5').numFmt = '#,##0';
  
  ws2.getCell('A6').value = "Simulasi Penerimaan PBB 2026";
  ws2.getCell('B6').value = "Sektor Lainnya";
  ws2.getCell('C6').value = { formula: "SUMIF('1. Hasil'!C2:C10000,\"Sektor Lainnya\",'1. Hasil'!AY2:AY10000)" };
  ws2.getCell('C6').numFmt = '#,##0';
  
  // Row 7
  ws2.getCell('A7').value = "Simulasi Penerimaan PBB 2026 (Collection Rate 100%)";
  ws2.getCell('B7').value = { formula: '(COUNT(\'1. Hasil\'!A2:A10000))&" NOP"' };
  ws2.getCell('C7').value = { formula: "SUM(C2:C6)" };
  ws2.getCell('C7').numFmt = '#,##0';
  
  // Row 8
  ws2.getCell('A8').value = "Target Penerimaan PBB 2026";
  ws2.getCell('C8').value = 110289165592;
  ws2.getCell('C8').numFmt = '#,##0';
  
  // Row 9
  ws2.getCell('A9').value = "Selisih antara Simulasi (Collection Rate 100%) & Target";
  ws2.getCell('C9').value = { formula: "C7-C8" };
  ws2.getCell('C9').numFmt = '#,##0';
  ws2.getCell('D9').value = { formula: 'IF(C9>0,"Tercapai","Tidak Tercapai")' };
  
  // Row 10 - B10 Format Persen 95%
  ws2.getCell('A10').value = { formula: '"Simulasi Penerimaan PBB 2026 (Collection Rate "&B10*100&"%)"' };
  ws2.getCell('B10').value = 0.95;
  ws2.getCell('B10').numFmt = '0%';
  ws2.getCell('C10').value = { formula: "C7*B10" };
  ws2.getCell('C10').numFmt = '#,##0';
  
  // Row 11
  ws2.getCell('A11').value = { formula: '"Selisih antara Simulasi (Collection Rate "&B10*100&"%)"&" Target"' };
  ws2.getCell('C11').value = { formula: "C10-C8" };
  ws2.getCell('C11').numFmt = '#,##0';
  ws2.getCell('D11').value = { formula: 'IF(C11>0,"Tercapai","Tidak Tercapai")' };
  
  // Row 13
  ws2.getCell('A13').value = "Poin";
  ws2.getCell('B13').value = { formula: '"Keterangan (BIT + "&E2*100&"% dan NDT + "&E14*100&"%)"' };
  ws2.getCell('C13').value = "Nilai";
  ws2.getCell('D13').value = "Keterangan";
  ws2.getCell('E13').value = "Skenario Kenaikan NDT";
  
  // Rows 14-18
  ws2.getCell('A14').value = { formula: "=A2" };
  ws2.getCell('B14').value = { formula: "=B2" };
  ws2.getCell('C14').value = { formula: "SUMIF('1. Hasil'!C2:C10000,\"Perkebunan\",'1. Hasil'!BC2:BC10000)" };
  ws2.getCell('C14').numFmt = '#,##0';
  
  // E14 - Format Persen 46%
  ws2.getCell('E14').value = 0.46;
  ws2.getCell('E14').numFmt = '0%';
  
  ws2.getCell('A15').value = { formula: "=A3" };
  ws2.getCell('B15').value = { formula: "=B3" };
  ws2.getCell('C15').value = { formula: "SUMIF('1. Hasil'!C2:C10000,\"Minerba\",'1. Hasil'!BC2:BC10000)" };
  ws2.getCell('C15').numFmt = '#,##0';
  
  ws2.getCell('A16').value = { formula: "=A4" };
  ws2.getCell('B16').value = { formula: "=B4" };
  ws2.getCell('C16').value = { formula: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (HTI)\",'1. Hasil'!BC2:BC10000)" };
  ws2.getCell('C16').numFmt = '#,##0';
  
  ws2.getCell('A17').value = { formula: "=A5" };
  ws2.getCell('B17').value = { formula: "=B5" };
  ws2.getCell('C17').value = { formula: "SUMIF('1. Hasil'!C2:C10000,\"Perhutanan (Hutan Alam)\",'1. Hasil'!BC2:BC10000)" };
  ws2.getCell('C17').numFmt = '#,##0';
  
  ws2.getCell('A18').value = { formula: "=A6" };
  ws2.getCell('B18').value = { formula: "=B6" };
  ws2.getCell('C18').value = { formula: "SUMIF('1. Hasil'!C2:C10000,\"Sektor Lainnya\",'1. Hasil'!BC2:BC10000)" };
  ws2.getCell('C18').numFmt = '#,##0';
  
  // Rows 19-23
  ws2.getCell('A19').value = { formula: "=A7" };
  ws2.getCell('B19').value = { formula: "=B7" };
  ws2.getCell('C19').value = { formula: "SUM(C14:C18)" };
  ws2.getCell('C19').numFmt = '#,##0';
  
  ws2.getCell('A20').value = { formula: "=A8" };
  ws2.getCell('C20').value = { formula: "=C8" };
  ws2.getCell('C20').numFmt = '#,##0';
  
  ws2.getCell('A21').value = { formula: "=A9" };
  ws2.getCell('C21').value = { formula: "C19-C20" };
  ws2.getCell('C21').numFmt = '#,##0';
  ws2.getCell('D21').value = { formula: 'IF(C21>0,"Tercapai","Tidak Tercapai")' };
  
  // Row 22 - B22 Format Persen 95%
  ws2.getCell('A22').value = { formula: '"Simulasi Penerimaan PBB 2026 (Collection Rate "&B22*100&"%)"' };
  ws2.getCell('B22').value = 0.95;
  ws2.getCell('B22').numFmt = '0%';
  ws2.getCell('C22').value = { formula: "C19*B22" };
  ws2.getCell('C22').numFmt = '#,##0';
  
  ws2.getCell('A23').value = { formula: '"Selisih antara Simulasi (Collection Rate "&B22*100&"%)"&" Target"' };
  ws2.getCell('C23').value = { formula: "C22-C20" };
  ws2.getCell('C23').numFmt = '#,##0';
  ws2.getCell('D23').value = { formula: 'IF(C23>0,"Tercapai","Tidak Tercapai")' };

  return workbook;
}
