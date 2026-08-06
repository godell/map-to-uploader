import React, { useState, useMemo } from "react";
import "./App.css";
import { 
  FileSpreadsheet, Upload, Download, Copy, Check, Filter, 
  Layers, RefreshCw, Eye, ArrowRight, CheckCircle2, AlertCircle,
  Database, HelpCircle, Sparkles, SlidersHorizontal, Table as TableIcon,
  Search
} from "lucide-react";
import { Button } from "./components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "./components/ui/card";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "./components/ui/tabs";
import { Badge } from "./components/ui/badge";
import { Input } from "./components/ui/input";
import { Label } from "./components/ui/label";
import { toast } from "sonner";

// Initial mock dataset from the user's uploaded Sourcedb.xlsx
const INITIAL_RAW_DATA = [
  {
    id: 1,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330146",
    "Transfer order item": "1",
    "Movement Type (WM)": "601",
    "Article": "SKE252175LKP005003",
    "Article Description": "SKE B SPORT SKECHERS JR YOUTH, BLACK",
    "Source target qty": "3",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-07-36A2",
    "Act.qty (dest)": "3",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298204"
  },
  {
    id: 2,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330146",
    "Transfer order item": "2",
    "Movement Type (WM)": "601",
    "Article": "SKE404203LZA00W003",
    "Article Description": "SKE B SP GLIDESTEP PLUS (K/M), WHITE",
    "Source target qty": "3",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-07-38A3",
    "Act.qty (dest)": "3",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298204"
  },
  {
    id: 3,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330146",
    "Transfer order item": "3",
    "Movement Type (WM)": "601",
    "Article": "SKE253021BKM005008",
    "Article Description": "SKE GO RUN SKX REIGN 2 (A/M) B, BLACK",
    "Source target qty": "8",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-08-34A1",
    "Act.qty (dest)": "8",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298204"
  },
  {
    id: 4,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330146",
    "Transfer order item": "4",
    "Movement Type (WM)": "601",
    "Article": "SKEJB112222B005T07",
    "Article Description": "SKE BOYS 6PK NT LOW CUT (, BLACK",
    "Source target qty": "7",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-10-03A4",
    "Act.qty (dest)": "7",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298204"
  },
  {
    id: 5,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330146",
    "Transfer order item": "5",
    "Movement Type (WM)": "601",
    "Article": "SKEBPT7680BL12W000",
    "Article Description": "SKE ATHLETIC BACKPACK (A/, BLUE",
    "Source target qty": "1",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-17-07A3",
    "Act.qty (dest)": "1",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298204"
  },
  {
    id: 6,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330147",
    "Transfer order item": "1",
    "Movement Type (WM)": "601",
    "Article": "SKE123456LKP005001",
    "Article Description": "SKECHERS RUNNING LITE, GREY",
    "Source target qty": "5",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-12-05B1",
    "Act.qty (dest)": "5",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298205"
  },
  {
    id: 7,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330147",
    "Transfer order item": "2",
    "Movement Type (WM)": "601",
    "Article": "SKE789012LZA00W002",
    "Article Description": "SKECHERS SPORT WALK, NAVY",
    "Source target qty": "4",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-15-22C4",
    "Act.qty (dest)": "4",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298205"
  },
  {
    id: 8,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330148",
    "Transfer order item": "1",
    "Movement Type (WM)": "601",
    "Article": "SKE998877BKM005003",
    "Article Description": "SKECHERS KIDSACTIVE, PINK",
    "Source target qty": "2",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-22-11D2",
    "Act.qty (dest)": "2",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298206"
  },
  {
    id: 9,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330149",
    "Transfer order item": "1",
    "Movement Type (WM)": "601",
    "Article": "SKE554433ABC005009",
    "Article Description": "SKECHERS MAX CUSHIONING, BLACK",
    "Source target qty": "6",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-30-01A1",
    "Act.qty (dest)": "6",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298207"
  },
  {
    id: 10,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330150",
    "Transfer order item": "1",
    "Movement Type (WM)": "601",
    "Article": "SKE112233XYZ005010",
    "Article Description": "SKECHERS SLIP-INS GO WALK, BLUE",
    "Source target qty": "10",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-45-08E3",
    "Act.qty (dest)": "10",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298208"
  },
  {
    id: 11,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330151",
    "Transfer order item": "1",
    "Movement Type (WM)": "601",
    "Article": "SKE991122QWE005011",
    "Article Description": "SKECHERS WORK RELAXED FIT, BROWN",
    "Source target qty": "3",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-55-12F2",
    "Act.qty (dest)": "3",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298209"
  },
  {
    id: 12,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330152",
    "Transfer order item": "1",
    "Movement Type (WM)": "601",
    "Article": "SKE774411ASD005012",
    "Article Description": "SKECHERS HI-LITES CLASSIC, WHITE",
    "Source target qty": "5",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-62-04G1",
    "Act.qty (dest)": "5",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298210"
  },
  {
    id: 13,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330153",
    "Transfer order item": "1",
    "Movement Type (WM)": "601",
    "Article": "SKE332211ZXC005013",
    "Article Description": "SKECHERS STREET CORSO, CHARCOAL",
    "Source target qty": "4",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-71-09H4",
    "Act.qty (dest)": "4",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298211"
  },
  {
    id: 14,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330154",
    "Transfer order item": "1",
    "Movement Type (WM)": "601",
    "Article": "SKE887766POI005014",
    "Article Description": "SKECHERS DLITES 4.0, MULTI",
    "Source target qty": "6",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-85-15J3",
    "Act.qty (dest)": "6",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298212"
  },
  {
    id: 15,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330155",
    "Transfer order item": "1",
    "Movement Type (WM)": "601",
    "Article": "SKE443322LKJ005015",
    "Article Description": "SKECHERS ARCH FIT REFINE, BLACK",
    "Source target qty": "8",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-92-02K1",
    "Act.qty (dest)": "8",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298213"
  },
  {
    id: 16,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330156",
    "Transfer order item": "1",
    "Movement Type (WM)": "601",
    "Article": "SKE665544MNB005016",
    "Article Description": "SKECHERS GO WALK FLEX, NAVY",
    "Source target qty": "9",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-98-11L2",
    "Act.qty (dest)": "9",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298214"
  },
  {
    id: 17,
    "Storage Location": "L001",
    "Warehouse Number": "W6A",
    "Transfer Order Number": "5330157",
    "Transfer order item": "1",
    "Movement Type (WM)": "601",
    "Article": "SKE101010VFR005017",
    "Article Description": "SKECHERS FLLEX COMP, BLACK/RED",
    "Source target qty": "2",
    "Source Storage Type": "G40",
    "Source Storage Bin": "M1-05-19M4",
    "Act.qty (dest)": "2",
    "Dest. Storage Type": "916",
    "Dest.Storage Bin": "8011298215"
  }
];

// Process a raw dataset row: extract 2 middle chars from Source Storage Bin starting at char 4
// e.g., "M1-07-36A2" -> Row "07"
const processDataset = (rawData) => {
  return rawData.map(item => {
    const bin = item["Source Storage Bin"] || "";
    let extractedRow = "XX";
    if (bin.length >= 5) {
      extractedRow = bin.substring(3, 5);
    } else if (bin.length >= 4) {
      extractedRow = bin.substring(3, bin.length);
    }
    return {
      ...item,
      "Row": extractedRow
    };
  });
};

export default function App() {
  const [data, setData] = useState(INITIAL_RAW_DATA);
  const [isGenerated, setIsGenerated] = useState(true);
  const [batchSize, setBatchSize] = useState(5);
  const [selectedCardRange, setSelectedCardRange] = useState(null); // e.g. { start: 1, end: 5 }
  const [searchQuery, setSearchQuery] = useState("");

  const processedData = useMemo(() => {
    return processDataset(data);
  }, [data]);

  // Parse Row string as numeric (e.g. "07" -> 7). Returns null if not numeric.
  const parseRowNumeric = (r) => {
    if (r === null || r === undefined) return null;
    const n = parseInt(r, 10);
    return Number.isNaN(n) ? null : n;
  };

  // Build map: Transfer Order Number -> Set of numeric Rows it contains
  // Also track raw row strings for display.
  const toRowInfo = useMemo(() => {
    const map = new Map(); // TO -> { numericRows: Set<number>, rawRows: Set<string> }
    processedData.forEach(item => {
      const to = item["Transfer Order Number"];
      if (!to) return;
      if (!map.has(to)) map.set(to, { numericRows: new Set(), rawRows: new Set() });
      const rec = map.get(to);
      const numeric = parseRowNumeric(item.Row);
      if (numeric !== null) rec.numericRows.add(numeric);
      if (item.Row) rec.rawRows.add(item.Row);
    });
    return map;
  }, [processedData]);

  // Max numeric Row present in the dataset (used to cap card ranges)
  const maxRow = useMemo(() => {
    let max = 0;
    toRowInfo.forEach(({ numericRows }) => {
      numericRows.forEach(n => { if (n > max) max = n; });
    });
    return max;
  }, [toRowInfo]);

  // Generate numeric range card batches from 1 .. maxRow (inclusive) with step = batchSize.
  // Example maxRow=36, batchSize=5 -> [1-5, 6-10, 11-15, 16-20, 21-25, 26-30, 31-35, 36-40]
  const cardBatches = useMemo(() => {
    const batches = [];
    if (maxRow === 0) return batches;
    for (let start = 1; start <= maxRow; start += batchSize) {
      const end = start + batchSize - 1;
      batches.push({
        type: "range",
        batchIndex: batches.length + 1,
        start,
        end,
      });
    }
    return batches;
  }, [maxRow, batchSize]);

  // Assign each Transfer Order to a card:
  //  - If ALL its numeric rows fall within a SINGLE card's [start, end] range, it belongs to that card.
  //  - Otherwise (rows span multiple cards, or has no numeric rows, or falls outside all cards), it goes to Mixed.
  const batchAssignments = useMemo(() => {
    const perBatch = new Map(); // batchIndex -> string[] of TO
    cardBatches.forEach(b => perBatch.set(b.batchIndex, []));
    const mixed = [];

    toRowInfo.forEach((info, to) => {
      const rows = Array.from(info.numericRows);
      if (rows.length === 0) {
        mixed.push(to);
        return;
      }
      const matchedBatches = new Set();
      let hasOutOfRange = false;
      rows.forEach(r => {
        const b = cardBatches.find(cb => r >= cb.start && r <= cb.end);
        if (b) matchedBatches.add(b.batchIndex);
        else hasOutOfRange = true;
      });
      if (!hasOutOfRange && matchedBatches.size === 1) {
        const [idx] = matchedBatches;
        perBatch.get(idx).push(to);
      } else {
        mixed.push(to);
      }
    });
    return { perBatch, mixed };
  }, [toRowInfo, cardBatches]);

  // Unique TOs to display in the table based on selection + search
  const uniqueTransferOrders = useMemo(() => {
    let list;
    if (!selectedCardRange) {
      list = Array.from(toRowInfo.keys());
    } else if (selectedCardRange.type === "mixed") {
      list = batchAssignments.mixed;
    } else {
      list = batchAssignments.perBatch.get(selectedCardRange.batchIndex) || [];
    }

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      list = list.filter(to => String(to).toLowerCase().includes(q));
    }
    return [...list].sort();
  }, [selectedCardRange, toRowInfo, batchAssignments, searchQuery]);

  // Helper: get sorted numeric rows for a TO (used in Mixed table)
  const getRowsForTO = (to) => {
    const info = toRowInfo.get(to);
    if (!info) return [];
    return Array.from(info.numericRows).sort((a, b) => a - b);
  };

  // Handle file upload simulation / CSV parsing
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const text = event.target.result;
        const lines = text.split("\n").filter(l => l.trim() !== "");
        if (lines.length < 2) {
          toast.error("File CSV tidak memiliki data yang cukup.");
          return;
        }
        const headers = lines[0].split(",").map(h => h.trim().replace(/^["']|["']$/g, ""));
        const parsedRows = [];

        for (let i = 1; i < lines.length; i++) {
          const currentLine = lines[i].split(",").map(val => val.trim().replace(/^["']|["']$/g, ""));
          if (currentLine.length >= headers.length) {
            const rowObj = { id: i };
            headers.forEach((hdr, idx) => {
              rowObj[hdr] = currentLine[idx] || "";
            });
            parsedRows.push(rowObj);
          }
        }

        if (parsedRows.length > 0) {
          setData(parsedRows);
          setIsGenerated(true);
          setSelectedCardRange(null);
          toast.success(`Berhasil memuat ${parsedRows.length} baris data dari file!`);
        } else {
          toast.error("Gagal memparsing baris CSV. Pastikan format benar.");
        }
      } catch (err) {
        console.error(err);
        toast.error("Terjadi kesalahan saat membaca file.");
      }
    };
    reader.readAsText(file);
  };

  // SAP TSV Copy helper — copies ONLY unique Transfer Order Numbers (1 column)
  const handleCopyToSAP = () => {
    if (uniqueTransferOrders.length === 0) {
      toast.error("Tidak ada Transfer Order Number untuk disalin.");
      return;
    }

    // Single column, newline-separated — siap paste ke kolom SAP GUI
    const tsvContent = uniqueTransferOrders.join("\n");

    navigator.clipboard.writeText(tsvContent).then(() => {
      toast.success(`${uniqueTransferOrders.length} Transfer Order Number unik disalin! Paste langsung (Ctrl+V) ke kolom SAP GUI.`);
    }).catch(() => {
      toast.error("Gagal menyalin ke clipboard.");
    });
  };

  const handleResetToDefault = () => {
    setData(INITIAL_RAW_DATA);
    setIsGenerated(true);
    setSelectedCardRange(null);
    setSearchQuery("");
    toast.info("Data dikembalikan ke sampel bawaan Sourcedb.xlsx");
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-teal-50/30 to-rose-50/20 text-slate-900 pb-20">
      {/* Header Banner */}
      <header className="border-b bg-white/80 backdrop-blur-md sticky top-0 z-30 shadow-xs">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="bg-teal-600 text-white p-2.5 rounded-xl shadow-md flex items-center justify-center">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <h1 className="font-bold text-lg tracking-tight bg-gradient-to-r from-teal-700 to-indigo-700 bg-clip-text text-transparent">
                SAP WM Transfer Order Processor
              </h1>
              <p className="text-xs text-slate-500 font-medium">
                Automated Row Extraction & SAP Grid Integration
              </p>
            </div>
          </div>

          <div className="flex items-center space-x-3">
            <Button
              variant="outline"
              size="sm"
              onClick={handleResetToDefault}
              data-testid="reset-data-btn"
              className="text-xs font-medium border-slate-200 hover:bg-slate-50"
            >
              <RefreshCw className="w-3.5 h-3.5 mr-1.5 text-teal-600" />
              Reset Data Sampel
            </Button>
            <div className="hidden sm:flex items-center bg-teal-50 text-teal-800 text-xs px-3 py-1.5 rounded-full font-semibold border border-teal-200">
              <Sparkles className="w-3.5 h-3.5 mr-1.5 text-teal-600" />
              {processedData.length} Records Loaded
            </div>
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 pt-8 space-y-8">
        
        {/* Step 1: Upload & Generator Control */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          
          {/* Upload Card */}
          <Card className="border-slate-200/80 shadow-sm bg-white/90 backdrop-blur-sm lg:col-span-1">
            <CardHeader className="pb-4">
              <div className="flex items-center justify-between">
                <CardTitle className="text-base font-semibold text-slate-800 flex items-center">
                  <Upload className="w-4 h-4 mr-2 text-teal-600" />
                  1. Upload Data Sumber
                </CardTitle>
                <Badge variant="secondary" className="bg-teal-100 text-teal-800 text-xs">
                  XLSX / CSV
                </Badge>
              </div>
              <CardDescription className="text-xs">
                Upload file data SAP Sourcedb Anda untuk diproses otomatis.
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="border-2 border-dashed border-slate-200 hover:border-teal-400 rounded-xl p-6 text-center transition-all bg-slate-50/50 group relative cursor-pointer">
                <input
                  type="file"
                  accept=".csv,.xlsx,.xls"
                  onChange={handleFileUpload}
                  className="absolute inset-0 opacity-0 cursor-pointer w-full h-full"
                  data-testid="file-upload-input"
                />
                <div className="flex flex-col items-center space-y-2">
                  <div className="p-3 bg-white shadow-xs rounded-full group-hover:scale-105 transition-transform text-teal-600">
                    <Database className="w-6 h-6" />
                  </div>
                  <div className="text-xs font-semibold text-slate-700">
                    Klik atau Seret file ke sini
                  </div>
                  <p className="text-[11px] text-slate-400">
                    Mendukung format .csv atau .xlsx (Sourcedb)
                  </p>
                </div>
              </div>

              <div className="pt-2 border-t border-slate-100 flex items-center justify-between text-xs text-slate-500">
                <span>Status Generator:</span>
                <span className="font-semibold text-emerald-600 flex items-center">
                  <CheckCircle2 className="w-3.5 h-3.5 mr-1" />
                  Kolom Row Aktif
                </span>
              </div>
            </CardContent>
          </Card>

          {/* Configuration & Rule Info Card */}
          <Card className="border-slate-200/80 shadow-sm bg-white/90 backdrop-blur-sm lg:col-span-2">
            <CardHeader className="pb-4">
              <div className="flex items-center justify-between">
                <CardTitle className="text-base font-semibold text-slate-800 flex items-center">
                  <SlidersHorizontal className="w-4 h-4 mr-2 text-indigo-600" />
                  2. Aturan Generator Kolom & Pengaturan Card
                </CardTitle>
                <Badge variant="outline" className="text-xs border-indigo-200 text-indigo-700 bg-indigo-50">
                  Automated Logic
                </Badge>
              </div>
              <CardDescription className="text-xs">
                Sistem otomatis menambahkan kolom <strong>Row</strong> dengan mengambil 2 karakter tengah dari kolom <strong>Source Storage Bin</strong> mulai dari karakter ke-4.
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                
                {/* Rule explanation box */}
                <div className="bg-indigo-50/60 border border-indigo-100 rounded-xl p-3.5 space-y-2">
                  <div className="text-xs font-semibold text-indigo-900 flex items-center">
                    <HelpCircle className="w-3.5 h-3.5 mr-1.5 text-indigo-600" />
                    Contoh Ekstraksi Bin:
                  </div>
                  <div className="text-xs text-indigo-700 space-y-1 font-mono bg-white/80 p-2 rounded-lg border border-indigo-100">
                    <div>Source Bin: <span className="text-rose-600 font-bold">M1-07-36A2</span></div>
                    <div>Karakter ke-4 & 2 berikutnya: <span className="text-emerald-600 font-bold">07</span></div>
                    <div>Hasil Kolom Row: <span className="text-indigo-600 font-bold">07</span></div>
                  </div>
                </div>

                {/* Batch Size Selector */}
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 flex flex-col justify-between space-y-3">
                  <div className="flex items-center justify-between">
                    <Label htmlFor="batch-select" className="text-xs font-semibold text-slate-700 flex items-center">
                      <Layers className="w-3.5 h-3.5 mr-1.5 text-teal-600" />
                      Ukuran Card Range per Grup:
                    </Label>
                    <span className="text-xs font-bold bg-teal-100 text-teal-800 px-2 py-0.5 rounded-md">
                      {batchSize} Row / Card
                    </span>
                  </div>
                  <div className="flex items-center space-x-2">
                    {[3, 5, 10, 15].map((size) => (
                      <Button
                        key={size}
                        size="sm"
                        variant={batchSize === size ? "default" : "outline"}
                        onClick={() => {
                          setBatchSize(size);
                          setSelectedCardRange(null);
                        }}
                        data-testid={`batch-size-${size}-btn`}
                        className={`flex-1 text-xs h-8 ${batchSize === size ? "bg-teal-600 hover:bg-teal-700 text-white" : "border-slate-200 hover:bg-slate-100 text-slate-700"}`}
                      >
                        {size} Row
                      </Button>
                    ))}
                  </div>
                  <p className="text-[11px] text-slate-400">
                    Membagi daftar Row unik menjadi card interaktif (Card 1, Card 2, dst).
                  </p>
                </div>

              </div>
            </CardContent>
          </Card>

        </div>

        {/* Step 3: Card Menu Section for Transfer Order Filtering */}
        <div className="space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-2 border-b border-slate-200">
            <div>
              <h2 className="text-lg font-bold text-slate-900 flex items-center">
                <Layers className="w-5 h-5 mr-2 text-teal-600" />
                3. Menu Card Range Row & Transfer Order Filter
              </h2>
              <p className="text-xs text-slate-500">
                Pilih salah satu card di bawah untuk memfilter Transfer Order Number dan menampilkan tabel siap-salin SAP.
              </p>
            </div>

            <div className="flex items-center space-x-2">
              <div className="relative">
                <Search className="absolute left-2.5 top-2.5 h-3.5 w-3.5 text-slate-400" />
                <Input
                  type="text"
                  placeholder="Cari data, TO, Artikel..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="pl-8 h-9 text-xs w-56 bg-white border-slate-200"
                  data-testid="search-input"
                />
              </div>
              {selectedCardRange && (
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={() => setSelectedCardRange(null)}
                  className="text-xs h-9 text-rose-600 hover:bg-rose-50 hover:text-rose-700"
                  data-testid="clear-filter-btn"
                >
                  Reset Filter
                </Button>
              )}
            </div>
          </div>

          {/* Cards Grid */}
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
            {cardBatches.map((batch) => {
              const isSelected = selectedCardRange && selectedCardRange.type === "range" && selectedCardRange.batchIndex === batch.batchIndex;
              const tosInBatch = batchAssignments.perBatch.get(batch.batchIndex) || [];

              return (
                <div
                  key={batch.batchIndex}
                  onClick={() => {
                    if (isSelected) {
                      setSelectedCardRange(null);
                    } else {
                      setSelectedCardRange(batch);
                    }
                  }}
                  data-testid={`card-batch-${batch.batchIndex}`}
                  className={`relative p-5 rounded-2xl transition-all cursor-pointer border shadow-xs flex flex-col justify-between group ${
                    isSelected
                      ? "bg-teal-50/90 border-teal-500 ring-2 ring-teal-400/30 shadow-md"
                      : "bg-white border-slate-200/80 hover:border-teal-300 hover:shadow-md hover:-translate-y-0.5"
                  }`}
                >
                  <div>
                    <div className="flex items-center justify-between mb-3">
                      <span className={`text-xs font-bold px-2.5 py-1 rounded-lg ${isSelected ? "bg-teal-600 text-white" : "bg-slate-100 text-slate-700 group-hover:bg-teal-100 group-hover:text-teal-800 transition-colors"}`}>
                        Card {batch.batchIndex}
                      </span>
                      <span className="text-[11px] font-medium text-slate-400">
                        Row {String(batch.start).padStart(2, "0")} - {String(batch.end).padStart(2, "0")}
                      </span>
                    </div>

                    <div className="space-y-1 mb-4">
                      <div className="text-xs font-semibold text-slate-700">Range Row:</div>
                      <div className="flex items-baseline gap-1.5">
                        <span className="text-2xl font-bold font-mono text-teal-700">
                          {String(batch.start).padStart(2, "0")}
                        </span>
                        <span className="text-slate-400 text-sm">→</span>
                        <span className="text-2xl font-bold font-mono text-teal-700">
                          {String(batch.end).padStart(2, "0")}
                        </span>
                      </div>
                    </div>
                  </div>

                  <div className="pt-3 border-t border-slate-100 flex items-center justify-between">
                    <div>
                      <div className="text-[10px] text-slate-400 uppercase tracking-wider font-semibold">Transfer Orders</div>
                      <div className="text-xs font-bold text-slate-800">{tosInBatch.length} TO Unik</div>
                    </div>
                    <div className={`p-2 rounded-full transition-all ${isSelected ? "bg-teal-600 text-white" : "bg-slate-50 text-slate-400 group-hover:bg-teal-50 group-hover:text-teal-600"}`}>
                      <ArrowRight className="w-3.5 h-3.5" />
                    </div>
                  </div>
                </div>
              );
            })}

            {/* Mixed Row Card — TOs whose rows span across multiple card ranges */}
            {cardBatches.length > 0 && (() => {
              const isSelected = selectedCardRange && selectedCardRange.type === "mixed";
              const mixedTOs = batchAssignments.mixed;
              return (
                <div
                  key="mixed-card"
                  onClick={() => {
                    if (isSelected) setSelectedCardRange(null);
                    else setSelectedCardRange({ type: "mixed", batchIndex: "mixed" });
                  }}
                  data-testid="card-batch-mixed"
                  className={`relative p-5 rounded-2xl transition-all cursor-pointer border shadow-xs flex flex-col justify-between group ${
                    isSelected
                      ? "bg-amber-50/90 border-amber-500 ring-2 ring-amber-400/30 shadow-md"
                      : "bg-white border-amber-200/80 hover:border-amber-400 hover:shadow-md hover:-translate-y-0.5"
                  }`}
                >
                  <div>
                    <div className="flex items-center justify-between mb-3">
                      <span className={`text-xs font-bold px-2.5 py-1 rounded-lg ${isSelected ? "bg-amber-600 text-white" : "bg-amber-100 text-amber-800 group-hover:bg-amber-200 transition-colors"}`}>
                        Mixed Row
                      </span>
                      <span className="text-[11px] font-medium text-amber-500 flex items-center">
                        <AlertCircle className="w-3 h-3 mr-1" />
                        Nyebar
                      </span>
                    </div>

                    <div className="space-y-1 mb-4">
                      <div className="text-xs font-semibold text-slate-700">TO Lintas Range:</div>
                      <div className="text-xs text-slate-500 leading-relaxed">
                        TO yang punya item di lebih dari 1 range card.
                      </div>
                    </div>
                  </div>

                  <div className="pt-3 border-t border-amber-100 flex items-center justify-between">
                    <div>
                      <div className="text-[10px] text-amber-500 uppercase tracking-wider font-semibold">Transfer Orders</div>
                      <div className="text-xs font-bold text-slate-800">{mixedTOs.length} TO Unik</div>
                    </div>
                    <div className={`p-2 rounded-full transition-all ${isSelected ? "bg-amber-600 text-white" : "bg-amber-50 text-amber-500 group-hover:bg-amber-100 group-hover:text-amber-700"}`}>
                      <ArrowRight className="w-3.5 h-3.5" />
                    </div>
                  </div>
                </div>
              );
            })()}
          </div>
        </div>

        {/* Step 4: SAP Transfer Order Table & Copy-Paste Section */}
        <Card className="border-slate-200/80 shadow-sm bg-white overflow-hidden" data-testid="sap-table-section">
          <CardHeader className="bg-slate-50/80 border-b border-slate-200 py-4 px-6">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <CardTitle className="text-base font-bold text-slate-800 flex items-center">
                  <TableIcon className="w-4 h-4 mr-2 text-teal-600" />
                  Daftar Transfer Order Number Unik ({uniqueTransferOrders.length} TO)
                </CardTitle>
                <CardDescription className="text-xs" data-testid="table-caption">
                  {(() => {
                    if (selectedCardRange && selectedCardRange.type === "range") {
                      return `Menampilkan TO yang SEMUA item Row-nya ada di Card ${selectedCardRange.batchIndex} (Range ${String(selectedCardRange.start).padStart(2, "0")} - ${String(selectedCardRange.end).padStart(2, "0")})`;
                    }
                    if (selectedCardRange && selectedCardRange.type === "mixed") {
                      return "Menampilkan TO Mixed Row — TO yang item Row-nya nyebar di lebih dari 1 range card";
                    }
                    if (searchQuery) {
                      return `Hasil pencarian untuk "${searchQuery}"`;
                    }
                    return "Menampilkan semua Transfer Order Number unik (Pilih card di atas untuk memfilter)";
                  })()}
                </CardDescription>
              </div>

              <div className="flex items-center space-x-3">
                <Button
                  onClick={handleCopyToSAP}
                  data-testid="copy-to-sap-btn"
                  className="bg-teal-600 hover:bg-teal-700 text-white text-xs font-semibold shadow-xs flex items-center"
                >
                  <Copy className="w-3.5 h-3.5 mr-1.5" />
                  Copy TO Unik ke SAP
                </Button>
              </div>
            </div>
          </CardHeader>
          <CardContent className="p-0">
            <div className="overflow-x-auto max-h-[500px]">
              <table className="w-full text-left border-collapse text-xs" data-testid="unique-to-table">
                <thead className="bg-slate-100/80 sticky top-0 z-10 text-slate-700 font-semibold border-b border-slate-200">
                  <tr>
                    <th className="py-3 px-4 w-20">No</th>
                    <th className="py-3 px-4 text-teal-800 font-bold bg-teal-50">Transfer Order Number</th>
                    {selectedCardRange && selectedCardRange.type === "mixed" && (
                      <th className="py-3 px-4 text-amber-800 font-bold bg-amber-50">Row Spread</th>
                    )}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {uniqueTransferOrders.length > 0 ? (
                    uniqueTransferOrders.map((to, idx) => {
                      const isMixedView = selectedCardRange && selectedCardRange.type === "mixed";
                      const rows = isMixedView ? getRowsForTO(to) : [];
                      return (
                        <tr key={to} data-testid={`to-row-${idx}`} className="hover:bg-slate-50/80 transition-colors">
                          <td className="py-2.5 px-4 text-slate-400 font-mono">{idx + 1}</td>
                          <td className="py-2.5 px-4 font-bold font-mono text-teal-800 text-sm">{to}</td>
                          {isMixedView && (
                            <td className="py-2.5 px-4 bg-amber-50/40">
                              <div className="flex flex-wrap gap-1">
                                {rows.map(r => (
                                  <span key={r} className="text-[11px] font-mono bg-amber-100 text-amber-800 px-1.5 py-0.5 rounded border border-amber-200">
                                    {String(r).padStart(2, "0")}
                                  </span>
                                ))}
                              </div>
                            </td>
                          )}
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan={selectedCardRange && selectedCardRange.type === "mixed" ? "3" : "2"} className="py-12 text-center text-slate-400">
                        Tidak ada Transfer Order Number yang sesuai dengan filter saat ini.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {/* Footer helper for SAP paste */}
            <div className="bg-slate-50 border-t border-slate-200 py-3 px-6 flex items-center justify-between text-xs text-slate-500">
              <div className="flex items-center space-x-2">
                <Check className="w-4 h-4 text-emerald-600" />
                <span>Tips SAP: Klik tombol <strong>Copy TO Unik</strong> di atas, lalu tekan <kbd className="bg-white px-1.5 py-0.5 rounded border border-slate-300 font-mono text-slate-700">Ctrl + V</kbd> pada kolom Transfer Order Number di SAP GUI.</span>
              </div>
              <div className="font-semibold text-slate-700">
                Total TO Unik: {uniqueTransferOrders.length}
              </div>
            </div>
          </CardContent>
        </Card>

      </main>
    </div>
  );
}
