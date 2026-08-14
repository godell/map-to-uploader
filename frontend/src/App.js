import { createClient } from '@supabase/supabase-js';
import React, { useState, useMemo, useEffect, useRef } from "react";
import "./App.css";
import { 
  FileSpreadsheet, Upload, Download, Copy, Check, Filter, 
  Layers, RefreshCw, Eye, ArrowRight, CheckCircle2, AlertCircle,
  Database, HelpCircle, Sparkles, SlidersHorizontal, Table as TableIcon,
  Search, Loader2, Printer, Package, Users, AlertTriangle, X,
  ClipboardList, CheckSquare, Calendar, Clock
} from "lucide-react";
import axios from "axios";
import * as XLSX from "xlsx";
import Barcode from "react-barcode";
import { QRCodeSVG } from "qrcode.react";
import { Button } from "./components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "./components/ui/card";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "./components/ui/tabs";
import { Badge } from "./components/ui/badge";
import { Input } from "./components/ui/input";
import { Label } from "./components/ui/label";
import { toast } from "sonner";

const SUPABASE_URL = process.env.REACT_APP_SUPABASE_URL;
const SUPABASE_ANON_KEY = process.env.REACT_APP_SUPABASE_ANON_KEY;
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

const BACKEND_URL = process.env.REACT_APP_BACKEND_URL;
const API = `${BACKEND_URL}/api`;

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
  const [selectedWing, setSelectedWing] = useState(null); 
  const [selectedSubCard, setSelectedSubCard] = useState(null); 
  const [searchQuery, setSearchQuery] = useState("");
  const [uploadStatus, setUploadStatus] = useState({ active: false, phase: "", fileName: "", rowsParsed: 0 });
  const [mixedSubFilter, setMixedSubFilter] = useState("all"); 
  const [copyPulse, setCopyPulse] = useState(false); 

  const [batchPickingOpen, setBatchPickingOpen] = useState(false); 
  const [batchSize, setBatchSize] = useState(null); 
  const [customBatchSize, setCustomBatchSize] = useState("");
  const [selectedBatchNumber, setSelectedBatchNumber] = useState(null); 
  const [pickerCount, setPickerCount] = useState(3);
  const [printModalOpen, setPrintModalOpen] = useState(false);
  const [printWarning, setPrintWarning] = useState(null); 
  const [isPrinting, setIsPrinting] = useState(false);
  const printingRef = useRef(false); 
  
  const [reprintCount, setReprintCount] = useState(0);
  const [printType, setPrintType] = useState('ptf');

  const processedData = useMemo(() => {
    return processDataset(data);
  }, [data]);

  const parseRowNumeric = (r) => {
    if (r === null || r === undefined) return null;
    const n = parseInt(r, 10);
    return Number.isNaN(n) ? null : n;
  };

  const parseQty = (q) => {
    if (q === null || q === undefined || q === "") return 0;
    const n = typeof q === "number" ? q : parseFloat(String(q).replace(/,/g, ""));
    return Number.isNaN(n) ? 0 : n;
  };

  const toRowInfo = useMemo(() => {
    const map = new Map();
    processedData.forEach(item => {
      const to = item["Transfer Order Number"];
      if (!to) return;
      if (!map.has(to)) {
        map.set(to, { numericRows: new Set(), rawRows: new Set(), bins: new Set(), articles: new Set(), qty: 0 });
      }
      const rec = map.get(to);
      const numeric = parseRowNumeric(item.Row);
      if (numeric !== null) rec.numericRows.add(numeric);
      if (item.Row) rec.rawRows.add(item.Row);
      if (item["Source Storage Bin"]) rec.bins.add(item["Source Storage Bin"]);
      if (item["Article"]) rec.articles.add(item["Article"]);
      rec.qty += parseQty(item["Source target qty"]);
    });
    return map;
  }, [processedData]);

  const datasetSummary = useMemo(() => {
    const articles = new Set();
    const rows = new Set();
    const bins = new Set();
    let totalQty = 0;
    processedData.forEach(item => {
      if (item["Article"]) articles.add(item["Article"]);
      if (item["Row"]) rows.add(item["Row"]);
      if (item["Source Storage Bin"]) bins.add(item["Source Storage Bin"]);
      totalQty += parseQty(item["Source target qty"]);
    });
    return {
      totalTO: toRowInfo.size,
      totalArticle: articles.size,
      totalRow: rows.size,
      totalBin: bins.size,
      totalQty,
    };
  }, [processedData, toRowInfo]);

  const WING_STRUCTURE = useMemo(() => ({
    left: {
      key: "left",
      label: "Wing Kiri",
      shortLabel: "Kiri",
      range: [1, 18],
      subRanges: [
        { key: "1-6", start: 1, end: 6 },
        { key: "7-12", start: 7, end: 12 },
        { key: "13-18", start: 13, end: 18 },
      ],
    },
    right: {
      key: "right",
      label: "Wing Kanan",
      shortLabel: "Kanan",
      range: [19, 36],
      subRanges: [
        { key: "19-24", start: 19, end: 24 },
        { key: "25-30", start: 25, end: 30 },
        { key: "31-36", start: 31, end: 36 },
      ],
    },
  }), []);

  const wingAssignments = useMemo(() => {
    const emptyRanges = (subs) => {
      const o = {};
      subs.forEach(sr => { o[sr.key] = []; });
      return o;
    };
    const result = {
      left: { ranges: emptyRanges(WING_STRUCTURE.left.subRanges), mixed: [] },
      right: { ranges: emptyRanges(WING_STRUCTURE.right.subRanges), mixed: [] },
      cross: [],
    };

    toRowInfo.forEach((info, to) => {
      const rows = Array.from(info.numericRows);
      if (rows.length === 0) {
        result.cross.push(to);
        return;
      }

      const inLeft = rows.some(r => r >= 1 && r <= 18);
      const inRight = rows.some(r => r >= 19 && r <= 36);
      const outOfBounds = rows.some(r => r < 1 || r > 36);

      if ((inLeft && inRight) || outOfBounds) {
        result.cross.push(to);
        return;
      }

      const wing = inLeft ? "left" : "right";
      const subs = WING_STRUCTURE[wing].subRanges;
      const matched = new Set();
      rows.forEach(r => {
        const s = subs.find(sr => r >= sr.start && r <= sr.end);
        if (s) matched.add(s.key);
      });

      if (matched.size === 1) {
        const [key] = matched;
        result[wing].ranges[key].push(to);
      } else {
        result[wing].mixed.push(to);
      }
    });

    return result;
  }, [toRowInfo, WING_STRUCTURE]);

  const computeBucketStats = (toList) => {
    let totalQty = 0;
    let totalBinCount = 0;
    toList.forEach(to => {
      const info = toRowInfo.get(to);
      if (!info) return;
      totalQty += info.qty;
      totalBinCount += info.bins.size;
    });
    const toCount = toList.length;
    return {
      toCount,
      totalQty,
      totalBinCount,
      avgBinPerTO: toCount > 0 ? Math.round(totalBinCount / toCount) : 0,
      avgQtyPerBin: totalBinCount > 0 ? (totalQty / totalBinCount) : 0,
    };
  };

  const wingLevelStats = useMemo(() => {
    const leftTOs = [
      ...Object.values(wingAssignments.left.ranges).flat(),
      ...wingAssignments.left.mixed,
    ];
    const rightTOs = [
      ...Object.values(wingAssignments.right.ranges).flat(),
      ...wingAssignments.right.mixed,
    ];
    return {
      left: computeBucketStats(leftTOs),
      right: computeBucketStats(rightTOs),
      cross: computeBucketStats(wingAssignments.cross),
    };
  }, [wingAssignments, toRowInfo]);

  const crossWingTOsSorted = useMemo(() => {
    return [...wingAssignments.cross].sort();
  }, [wingAssignments.cross]);

  const batches = useMemo(() => {
    if (!batchSize || batchSize <= 0) return [];
    const result = [];
    let running = 1;
    for (let i = 0; i < crossWingTOsSorted.length; i += batchSize) {
      const chunk = crossWingTOsSorted.slice(i, i + batchSize);
      const codedTOs = chunk.map((to, idx) => ({ code: running + idx, to }));
      running += chunk.length;
      const stats = computeBucketStats(chunk);
      const randomSuffix = Math.floor(100 + Math.random() * 900); 

      result.push({
        batchNumber: result.length + 1,
        tos: chunk,
        codedTOs,
        stats,
        randomSuffix 
      });
    }
    return result;
  }, [crossWingTOsSorted, batchSize]);

  const selectedBatch = useMemo(() => {
    if (!selectedBatchNumber) return null;
    return batches.find(b => b.batchNumber === selectedBatchNumber) || null;
  }, [batches, selectedBatchNumber]);

  const buildRowPages = (batch) => {
    if (!batch) return [];

    const ITEMS_PER_PAGE = 10;
    const toSet = new Set(batch.tos);

    const items = processedData.filter(item =>
      toSet.has(item["Transfer Order Number"])
    );

    const groups = new Map();

    items.forEach(it => {
      const row = it.Row || "??";
      if (!groups.has(row)) {
        groups.set(row, []);
      }
      groups.get(row).push(it);
    });

    const rows = Array.from(groups.keys()).sort((a, b) => {
      const na = parseInt(a, 10);
      const nb = parseInt(b, 10);

      if (Number.isNaN(na) && Number.isNaN(nb)) {
        return String(a).localeCompare(String(b));
      }

      if (Number.isNaN(na)) return 1;
      if (Number.isNaN(nb)) return -1;

      return na - nb;
    });

    const result = [];

    rows.forEach(rowKey => {
      const rowItems = [...groups.get(rowKey)].sort((a, b) =>
        String(a["Source Storage Bin"] || "")
          .localeCompare(
            String(b["Source Storage Bin"] || "")
          )
      );

      const toUniq = new Set();
      const artUniq = new Set();
      let qtySum = 0;

      rowItems.forEach(it => {
        if (it["Transfer Order Number"]) {
          toUniq.add(it["Transfer Order Number"]);
        }

        if (it["Article"]) {
          artUniq.add(it["Article"]);
        }

        qtySum += parseQty(
          it["Source target qty"]
        );
      });

      const stats = {
        totalTO: toUniq.size,
        totalArticle: artUniq.size,
        totalQty: qtySum,
      };

      const totalRowPages = Math.max(
        1,
        Math.ceil(
          rowItems.length / ITEMS_PER_PAGE
        )
      );

      for (
        let start = 0;
        start < rowItems.length;
        start += ITEMS_PER_PAGE
      ) {
        const pageItems = rowItems.slice(
          start,
          start + ITEMS_PER_PAGE
        );

        const rowPageNumber = Math.floor(start / ITEMS_PER_PAGE) + 1;

        result.push({
          rowKey,
          items: pageItems,
          stats,
          totalRowItems: rowItems.length,
          rowPageNumber,
          rowTotalPages: totalRowPages,
          isFirstPage: rowPageNumber === 1,
        });
      }
    });

    return result;
  };

  const currentTOList = useMemo(() => {
    if (!selectedWing) return Array.from(toRowInfo.keys());
    if (selectedWing === "cross") {
      if (selectedBatch) return selectedBatch.tos;
      return wingAssignments.cross;
    }

    const wingBucket = wingAssignments[selectedWing];
    if (!selectedSubCard) {
      return [
        ...Object.values(wingBucket.ranges).flat(),
        ...wingBucket.mixed,
      ];
    }
    if (selectedSubCard.type === "range") {
      return wingBucket.ranges[selectedSubCard.key] || [];
    }
    if (selectedSubCard.type === "mixed") {
      const base = wingBucket.mixed;
      if (mixedSubFilter === "all") return base;
      if (mixedSubFilter === ">5") {
        return base.filter(to => {
          const info = toRowInfo.get(to);
          return info && info.numericRows.size > 5;
        });
      }
      const target = parseInt(mixedSubFilter, 10);
      return base.filter(to => {
        const info = toRowInfo.get(to);
        return info && info.numericRows.size === target;
      });
    }
    return [];
  }, [selectedWing, selectedSubCard, wingAssignments, mixedSubFilter, toRowInfo, selectedBatch]);

  const uniqueTransferOrders = useMemo(() => {
    let list = currentTOList;
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      list = list.filter(to => String(to).toLowerCase().includes(q));
    }
    return [...list].sort();
  }, [currentTOList, searchQuery]);

  const isFiltered = selectedWing !== null;

  const formatBatchCode = (batchNumber, randomSuffix) => {
    const d = new Date();
    const yy = String(d.getFullYear()).slice(-2);
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    const suffixStr = String(randomSuffix || "000").padStart(3, '0');
    return `PTF-Batch-${batchNumber}-${yy}${mm}${dd}-${suffixStr}`;
  };

  const preparePrintOrientation = (orientation) => {
    const oldStyle = document.getElementById('dynamic-print-page');

    if (oldStyle) {
      oldStyle.remove();
    }

    const style = document.createElement('style');

    style.id = 'dynamic-print-page';

    style.textContent = `
      @media print {
        @page {
          size: A4 ${orientation};
          margin: 10mm;
        }
      }
    `;

    document.head.appendChild(style);
  };

  const handleRequestPrint = async () => {
    if (!selectedBatch) return;

    setPrintType('ptf');

    try {
      const toNumbers = selectedBatch.tos;

      const { data: prev, error } = await supabase
        .from('print_history')
        .select('*')
        .in('to_number', toNumbers);

      if (error) throw error;

      if (prev && prev.length > 0) {
        const counts = {};

        prev.forEach(p => {
          counts[p.to_number] = (counts[p.to_number] || 0) + 1;
        });

        const maxReprint = Math.max(...Object.values(counts));

        setReprintCount(maxReprint);
        setPrintWarning({ previousPrints: prev });
        setPrintModalOpen(true);

      } else {
        setReprintCount(0);
        setPrintWarning(null);
        setPrintModalOpen(false);

        proceedWithPrint();
      }

    } catch (err) {
      console.error(err);

      toast.error(
        "Gagal memeriksa riwayat print. Melanjutkan tanpa cek duplikat."
      );

      setReprintCount(0);
      setPrintModalOpen(false);

      proceedWithPrint();
    }
  };

  const handlePrintChecklist = () => {
    if (!selectedBatch) return;

    setPrintType('checklist');

    document.body.classList.remove('print-ptf');
    document.body.classList.add('print-checklist');

    setTimeout(async () => {
      try {
        preparePrintOrientation('portrait');

        await new Promise(resolve => {
          requestAnimationFrame(() => {
            requestAnimationFrame(resolve);
          });
        });

        window.print();

      } finally {
        document.body.classList.remove('print-checklist');

        const printStyle = document.getElementById(
          'dynamic-print-page'
        );

        if (printStyle) {
          printStyle.remove();
        }
      }
    }, 300);
  };

  const proceedWithPrint = async () => {
    if (!selectedBatch) return;
    if (printingRef.current) return;

    printingRef.current = true;
    setPrintWarning(null);
    setIsPrinting(true);

    document.body.classList.remove('print-checklist');
    document.body.classList.add('print-ptf');

    setTimeout(async () => {
      try {
        preparePrintOrientation('landscape');

        await new Promise(resolve => {
          requestAnimationFrame(() => {
            requestAnimationFrame(resolve);
          });
        });

        window.print();

      } finally {
        document.body.classList.remove('print-ptf');

        const printStyle = document.getElementById(
          'dynamic-print-page'
        );

        if (printStyle) {
          printStyle.remove();
        }

        try {
          const recordsToInsert = selectedBatch.tos.map(to => ({
            to_number: to,
            batch_code: formatBatchCode(
              selectedBatch.batchNumber,
              selectedBatch.randomSuffix
            ),
          }));

          const { error } = await supabase
            .from('print_history')
            .insert(recordsToInsert);

          if (error) throw error;

          toast.success(
            `Pick ticket Batch ${selectedBatch.batchNumber} berhasil disimpan di database.`
          );

        } catch (e) {
          console.error(e);

          toast.error(
            "Print sukses, tetapi gagal menyimpan history ke Supabase."
          );
        }

        setIsPrinting(false);
        printingRef.current = false;
      }

    }, 200);
  };

  useEffect(() => {
    setMixedSubFilter("all");
  }, [selectedWing, selectedSubCard]);

  useEffect(() => {
    setBatchPickingOpen(false);
    setBatchSize(null);
    setCustomBatchSize("");
    setSelectedBatchNumber(null);
  }, [selectedWing]);

  useEffect(() => {
    setSelectedBatchNumber(null);
  }, [batchSize]);

  useEffect(() => {
    setSelectedSubCard(null);
  }, [selectedWing]);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setUploadStatus({ active: true, phase: "reading", fileName: file.name, rowsParsed: 0 });

    const isExcel = /\.(xlsx|xls)$/i.test(file.name);
    const reader = new FileReader();

    reader.onerror = () => {
      setUploadStatus({ active: false, phase: "", fileName: "", rowsParsed: 0 });
      toast.error("Gagal membaca file. Coba ulangi upload.");
    };

    reader.onload = (event) => {
      setTimeout(() => {
        try {
          setUploadStatus(s => ({ ...s, phase: "parsing" }));

          let parsedRows = [];

          if (isExcel) {
            const buffer = event.target.result;
            const workbook = XLSX.read(buffer, { type: "array" });
            const firstSheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[firstSheetName];
            const json = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
            parsedRows = json.map((row, i) => ({ id: i + 1, ...row }));
          } else {
            const text = event.target.result;
            const lines = text.split(/\r?\n/).filter(l => l.trim() !== "");
            if (lines.length < 2) {
              setUploadStatus({ active: false, phase: "", fileName: "", rowsParsed: 0 });
              toast.error("File CSV tidak memiliki data yang cukup.");
              return;
            }
            const headers = lines[0].split(",").map(h => h.trim().replace(/^["']|["']$/g, ""));
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
          }

          if (parsedRows.length === 0) {
            setUploadStatus({ active: false, phase: "", fileName: "", rowsParsed: 0 });
            toast.error("Gagal memparsing file. Pastikan sheet pertama berisi data dengan header di baris pertama.");
            return;
          }

          const sample = parsedRows[0];
          if (!("Source Storage Bin" in sample) || !("Transfer Order Number" in sample)) {
            setUploadStatus({ active: false, phase: "", fileName: "", rowsParsed: 0 });
            toast.error("Kolom wajib tidak ditemukan: 'Source Storage Bin' dan/atau 'Transfer Order Number'. Cek nama kolom di file.");
            return;
          }

          setUploadStatus(s => ({ ...s, phase: "generating", rowsParsed: parsedRows.length }));

          setTimeout(() => {
            setData(parsedRows);
            setIsGenerated(true);
            setSelectedWing(null);
            setSelectedSubCard(null);
            setSearchQuery("");
            setUploadStatus({ active: false, phase: "", fileName: "", rowsParsed: 0 });
            toast.success(`Berhasil memuat ${parsedRows.length.toLocaleString("id-ID")} baris data dari "${file.name}"!`);
          }, 50);
        } catch (err) {
          console.error(err);
          setUploadStatus({ active: false, phase: "", fileName: "", rowsParsed: 0 });
          toast.error(`Terjadi kesalahan saat memproses file: ${err.message || err}`);
        }
      }, 30);
    };

    if (isExcel) {
      reader.readAsArrayBuffer(file);
    } else {
      reader.readAsText(file);
    }
    e.target.value = "";
  };

  const handleCopyToSAP = () => {
    if (uniqueTransferOrders.length === 0) {
      toast.error("Tidak ada Transfer Order Number untuk disalin.");
      return;
    }

    const tsvContent = uniqueTransferOrders.join("\n");

    navigator.clipboard.writeText(tsvContent).then(() => {
      setCopyPulse(true);
      setTimeout(() => setCopyPulse(false), 1400);
      toast.success(
        `${uniqueTransferOrders.length.toLocaleString("id-ID")} Transfer Order Number disalin!`,
        {
          description: "Langsung tekan Ctrl + V di kolom Transfer Order Number pada SAP GUI.",
          duration: 3500,
          icon: <Check className="w-4 h-4 text-emerald-600" />,
        }
      );
    }).catch(() => {
      toast.error("Gagal menyalin ke clipboard.");
    });
  };

  const handleResetData = () => {
    setData([]);
    setIsGenerated(false);
    setSelectedWing(null);
    setSelectedSubCard(null);
    setSearchQuery("");
    toast.info("Data upload berhasil dihapus. Silakan upload file baru.");
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-teal-50/30 to-rose-50/20 text-slate-900 pb-20">
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
              onClick={handleResetData}
              data-testid="reset-data-btn"
              className="text-xs font-medium border-slate-200 hover:bg-slate-50"
            >
              <RefreshCw className="w-3.5 h-3.5 mr-1.5 text-teal-600" />
              Reset Data
            </Button>
            <div className="hidden sm:flex items-center bg-teal-50 text-teal-800 text-xs px-3 py-1.5 rounded-full font-semibold border border-teal-200">
              <Sparkles className="w-3.5 h-3.5 mr-1.5 text-teal-600" />
              {processedData.length} Records Loaded
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 pt-8 space-y-8">
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
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
              <div className="bg-indigo-50/60 border border-indigo-100 rounded-xl p-4 space-y-3" data-testid="dataset-summary">
                <div className="flex items-center justify-between">
                  <div className="text-xs font-semibold text-indigo-900 flex items-center">
                    <Sparkles className="w-3.5 h-3.5 mr-1.5 text-indigo-600" />
                    Ringkasan Dataset
                  </div>
                  <Badge variant="outline" className="text-[10px] border-indigo-200 text-indigo-700 bg-indigo-50">
                    5 Metrik Global
                  </Badge>
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
                  <div className="bg-white/85 border border-indigo-100 rounded-lg px-3 py-2" data-testid="summary-total-to">
                    <div className="text-[10px] text-slate-500 uppercase tracking-wider font-semibold">Total TO</div>
                    <div className="text-sm font-bold text-indigo-700 font-mono">{datasetSummary.totalTO.toLocaleString("id-ID")}</div>
                  </div>
                  <div className="bg-white/85 border border-indigo-100 rounded-lg px-3 py-2" data-testid="summary-total-article">
                    <div className="text-[10px] text-slate-500 uppercase tracking-wider font-semibold">Total Article</div>
                    <div className="text-sm font-bold text-indigo-700 font-mono">{datasetSummary.totalArticle.toLocaleString("id-ID")}</div>
                  </div>
                  <div className="bg-white/85 border border-indigo-100 rounded-lg px-3 py-2" data-testid="summary-total-row">
                    <div className="text-[10px] text-slate-500 uppercase tracking-wider font-semibold">Total Row</div>
                    <div className="text-sm font-bold text-indigo-700 font-mono">{datasetSummary.totalRow.toLocaleString("id-ID")}</div>
                  </div>
                  <div className="bg-white/85 border border-indigo-100 rounded-lg px-3 py-2" data-testid="summary-total-bin">
                    <div className="text-[10px] text-slate-500 uppercase tracking-wider font-semibold">Total Bin</div>
                    <div className="text-sm font-bold text-indigo-700 font-mono">{datasetSummary.totalBin.toLocaleString("id-ID")}</div>
                  </div>
                  <div className="bg-white/85 border border-emerald-100 rounded-lg px-3 py-2" data-testid="summary-total-qty">
                    <div className="text-[10px] text-emerald-700 uppercase tracking-wider font-semibold">Total Qty</div>
                    <div className="text-sm font-bold text-emerald-700 font-mono">{datasetSummary.totalQty.toLocaleString("id-ID")}</div>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-2 text-[11px]" data-testid="wing-structure-info">
                <div className="bg-teal-50/70 border border-teal-100 rounded-lg px-3 py-2">
                  <div className="text-teal-800 font-semibold">Wing Kiri</div>
                  <div className="text-teal-600 font-mono">Row 01 - 18 · sub 1-6 / 7-12 / 13-18</div>
                </div>
                <div className="bg-indigo-50/70 border border-indigo-100 rounded-lg px-3 py-2">
                  <div className="text-indigo-800 font-semibold">Wing Kanan</div>
                  <div className="text-indigo-600 font-mono">Row 19 - 36 · sub 19-24 / 25-30 / 31-36</div>
                </div>
                <div className="bg-rose-50/70 border border-rose-100 rounded-lg px-3 py-2">
                  <div className="text-rose-800 font-semibold">Cross Wing</div>
                  <div className="text-rose-600 font-mono">TO dengan Row di 2 wing</div>
                </div>
              </div>
            </CardContent>
          </Card>
        </div>

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
              {selectedWing && (
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={() => { setSelectedWing(null); setSelectedSubCard(null); }}
                  className="text-xs h-9 text-rose-600 hover:bg-rose-50 hover:text-rose-700"
                  data-testid="clear-filter-btn"
                >
                  Reset Filter
                </Button>
              )}
            </div>
          </div>

          {data.length === 0 ? (
            <div
              className="bg-white border border-dashed border-slate-300 rounded-2xl p-10 flex flex-col items-center justify-center text-center"
              data-testid="empty-state"
            >
              <div className="w-14 h-14 rounded-full bg-slate-50 border border-slate-200 flex items-center justify-center mb-3">
                <Database className="w-6 h-6 text-slate-400" />
              </div>
              <div className="text-sm font-semibold text-slate-700 mb-1">
                Belum ada data yang di-upload.
              </div>
              <div className="text-xs text-slate-400 max-w-md">
                Upload file <span className="font-mono">Sourcedb.xlsx</span> atau CSV di panel <strong>1. Upload Data Sumber</strong> di atas untuk mulai memfilter Transfer Order.
              </div>
            </div>
          ) : (
          <>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4" data-testid="wing-cards-grid">
            {[
              { key: "left", label: "Wing Kiri", range: "Row 01 - 18", color: "teal", stats: wingLevelStats.left },
              { key: "right", label: "Wing Kanan", range: "Row 19 - 36", color: "indigo", stats: wingLevelStats.right },
              { key: "cross", label: "Cross Wing", range: "TO Lintas Wing", color: "rose", stats: wingLevelStats.cross },
            ].map(wing => {
              const active = selectedWing === wing.key;
              const c = wing.color;
              const colorMap = {
                teal: {
                  activeBg: "bg-teal-50/90 border-teal-500 ring-2 ring-teal-400/30",
                  idleBg: "bg-white border-teal-200/80 hover:border-teal-400",
                  chipActive: "bg-teal-600 text-white",
                  chipIdle: "bg-teal-100 text-teal-800",
                  numColor: "text-teal-700",
                  metaColor: "text-teal-400",
                  arrowActive: "bg-teal-600 text-white",
                  arrowIdle: "bg-teal-50 text-teal-500",
                },
                indigo: {
                  activeBg: "bg-indigo-50/90 border-indigo-500 ring-2 ring-indigo-400/30",
                  idleBg: "bg-white border-indigo-200/80 hover:border-indigo-400",
                  chipActive: "bg-indigo-600 text-white",
                  chipIdle: "bg-indigo-100 text-indigo-800",
                  numColor: "text-indigo-700",
                  metaColor: "text-indigo-400",
                  arrowActive: "bg-indigo-600 text-white",
                  arrowIdle: "bg-indigo-50 text-indigo-500",
                },
                rose: {
                  activeBg: "bg-rose-50/90 border-rose-500 ring-2 ring-rose-400/30",
                  idleBg: "bg-white border-rose-200/80 hover:border-rose-400",
                  chipActive: "bg-rose-600 text-white",
                  chipIdle: "bg-rose-100 text-rose-800",
                  numColor: "text-rose-700",
                  metaColor: "text-rose-400",
                  arrowActive: "bg-rose-600 text-white",
                  arrowIdle: "bg-rose-50 text-rose-500",
                },
              }[c];
              return (
                <div
                  key={wing.key}
                  onClick={() => {
                    if (active) { setSelectedWing(null); setSelectedSubCard(null); }
                    else setSelectedWing(wing.key);
                  }}
                  data-testid={`wing-card-${wing.key}`}
                  className={`relative p-5 rounded-2xl transition-all cursor-pointer border shadow-xs flex flex-col justify-between group ${
                    active ? colorMap.activeBg + " shadow-md" : colorMap.idleBg + " hover:shadow-md hover:-translate-y-0.5"
                  }`}
                >
                  <div>
                    <div className="flex items-center justify-between mb-3">
                      <span className={`text-sm font-bold px-3 py-1 rounded-lg ${active ? colorMap.chipActive : colorMap.chipIdle}`}>
                        {wing.label}
                      </span>
                      <span className={`text-[11px] font-medium ${colorMap.metaColor}`}>{wing.range}</span>
                    </div>

                    <div className="space-y-1 mb-3">
                      <div className="text-xs font-semibold text-slate-700">Total TO:</div>
                      <div className={`text-3xl font-bold font-mono ${colorMap.numColor}`}>
                        {wing.stats.toCount.toLocaleString("id-ID")}
                      </div>
                    </div>

                    <div className="grid grid-cols-3 gap-1.5 mb-1" data-testid={`wing-metrics-${wing.key}`}>
                      <div className="bg-emerald-50/70 border border-emerald-100 rounded-md px-2 py-1">
                        <div className="text-[9px] uppercase font-semibold tracking-wider text-emerald-700">Total Qty</div>
                        <div className="text-xs font-bold font-mono text-emerald-800">{wing.stats.totalQty.toLocaleString("id-ID")}</div>
                      </div>
                      <div className="bg-indigo-50/70 border border-indigo-100 rounded-md px-2 py-1">
                        <div className="text-[9px] uppercase font-semibold tracking-wider text-indigo-700">Avg Bin/TO</div>
                        <div className="text-xs font-bold font-mono text-indigo-800">{wing.stats.avgBinPerTO}</div>
                      </div>
                      <div className="bg-amber-50/70 border border-amber-100 rounded-md px-2 py-1">
                        <div className="text-[9px] uppercase font-semibold tracking-wider text-amber-700">Avg Qty/Bin</div>
                        <div className="text-xs font-bold font-mono text-amber-800">{wing.stats.avgQtyPerBin.toFixed(1)}</div>
                      </div>
                    </div>
                  </div>

                  <div className="pt-3 border-t border-slate-100 flex items-center justify-between">
                    <div className="text-[10px] text-slate-500 uppercase tracking-wider font-semibold">
                      {active ? "Aktif · klik untuk tutup" : "Klik untuk buka sub-card"}
                    </div>
                    <div className={`p-2 rounded-full transition-all ${active ? colorMap.arrowActive : colorMap.arrowIdle + " group-hover:bg-opacity-80"}`}>
                      <ArrowRight className="w-3.5 h-3.5" />
                    </div>
                  </div>
                </div>
              );
            })}
          </div>

          {selectedWing && selectedWing !== "cross" && (() => {
            const wingDef = WING_STRUCTURE[selectedWing];
            const wingBucket = wingAssignments[selectedWing];
            const wingColor = selectedWing === "left" ? "teal" : "indigo";
            const activeChipBg = wingColor === "teal" ? "bg-teal-600 text-white" : "bg-indigo-600 text-white";
            const idleChipBg = wingColor === "teal" ? "bg-teal-100 text-teal-800" : "bg-indigo-100 text-indigo-800";
            const idleBorder = wingColor === "teal" ? "border-teal-200/80 hover:border-teal-400" : "border-indigo-200/80 hover:border-indigo-400";
            const activeRing = wingColor === "teal" ? "border-teal-500 ring-2 ring-teal-400/30 bg-teal-50/70" : "border-indigo-500 ring-2 ring-indigo-400/30 bg-indigo-50/70";
            const numColor = wingColor === "teal" ? "text-teal-700" : "text-indigo-700";
            return (
              <div
                className={`mt-2 p-4 rounded-2xl border ${wingColor === "teal" ? "bg-teal-50/40 border-teal-200/60" : "bg-indigo-50/40 border-indigo-200/60"}`}
                data-testid={`subcards-panel-${selectedWing}`}
              >
                <div className="text-xs font-semibold text-slate-700 mb-3 flex items-center">
                  <Layers className={`w-3.5 h-3.5 mr-1.5 ${wingColor === "teal" ? "text-teal-600" : "text-indigo-600"}`} />
                  Sub-card {wingDef.label} — pilih rentang atau Mixed:
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                  {wingDef.subRanges.map(sr => {
                    const tos = wingBucket.ranges[sr.key] || [];
                    const stats = computeBucketStats(tos);
                    const active = selectedSubCard && selectedSubCard.type === "range" && selectedSubCard.wing === selectedWing && selectedSubCard.key === sr.key;
                    return (
                      <div
                        key={sr.key}
                        onClick={() => {
                          if (active) setSelectedSubCard(null);
                          else setSelectedSubCard({ type: "range", wing: selectedWing, key: sr.key, start: sr.start, end: sr.end });
                        }}
                        data-testid={`subcard-range-${selectedWing}-${sr.key}`}
                        className={`p-4 rounded-xl border cursor-pointer transition-all bg-white ${
                          active ? activeRing + " shadow-md" : idleBorder + " hover:shadow-md hover:-translate-y-0.5"
                        }`}
                      >
                        <div className="flex items-center justify-between mb-2">
                          <span className={`text-[11px] font-bold px-2 py-0.5 rounded ${active ? activeChipBg : idleChipBg}`}>
                            {wingDef.shortLabel} {sr.key}
                          </span>
                          <span className="text-[10px] font-medium text-slate-400">Row {String(sr.start).padStart(2, "0")}-{String(sr.end).padStart(2, "0")}</span>
                        </div>
                        <div className={`text-2xl font-bold font-mono mb-2 ${numColor}`}>
                          {stats.toCount.toLocaleString("id-ID")}
                          <span className="text-[10px] text-slate-400 font-normal ml-1">TO Unik</span>
                        </div>
                        <div className="grid grid-cols-3 gap-1">
                          <div className="bg-emerald-50/70 border border-emerald-100 rounded px-1.5 py-1">
                            <div className="text-[8px] uppercase font-semibold tracking-wider text-emerald-700">Qty</div>
                            <div className="text-[11px] font-bold font-mono text-emerald-800">{stats.totalQty.toLocaleString("id-ID")}</div>
                          </div>
                          <div className="bg-indigo-50/70 border border-indigo-100 rounded px-1.5 py-1">
                            <div className="text-[8px] uppercase font-semibold tracking-wider text-indigo-700">Bin/TO</div>
                            <div className="text-[11px] font-bold font-mono text-indigo-800">{stats.avgBinPerTO}</div>
                          </div>
                          <div className="bg-amber-50/70 border border-amber-100 rounded px-1.5 py-1">
                            <div className="text-[8px] uppercase font-semibold tracking-wider text-amber-700">Qty/Bin</div>
                            <div className="text-[11px] font-bold font-mono text-amber-800">{stats.avgQtyPerBin.toFixed(1)}</div>
                          </div>
                        </div>
                      </div>
                    );
                  })}

                  {(() => {
                    const tos = wingBucket.mixed;
                    const stats = computeBucketStats(tos);
                    const active = selectedSubCard && selectedSubCard.type === "mixed" && selectedSubCard.wing === selectedWing;
                    return (
                      <div
                        onClick={() => {
                          if (active) setSelectedSubCard(null);
                          else setSelectedSubCard({ type: "mixed", wing: selectedWing, key: "mixed" });
                        }}
                        data-testid={`subcard-mixed-${selectedWing}`}
                        className={`p-4 rounded-xl border cursor-pointer transition-all ${
                          active
                            ? "bg-amber-50 border-amber-500 ring-2 ring-amber-400/30 shadow-md"
                            : "bg-white border-amber-200 hover:border-amber-400 hover:shadow-md hover:-translate-y-0.5"
                        }`}
                      >
                        <div className="flex items-center justify-between mb-2">
                          <span className={`text-[11px] font-bold px-2 py-0.5 rounded ${active ? "bg-amber-600 text-white" : "bg-amber-100 text-amber-800"}`}>
                            Mixed {wingDef.shortLabel}
                          </span>
                          <span className="text-[10px] font-medium text-amber-500 flex items-center">
                            <AlertCircle className="w-3 h-3 mr-0.5" />
                            Nyebar
                          </span>
                        </div>
                        <div className="text-2xl font-bold font-mono mb-2 text-amber-700">
                          {stats.toCount.toLocaleString("id-ID")}
                          <span className="text-[10px] text-slate-400 font-normal ml-1">TO Unik</span>
                        </div>
                        <div className="grid grid-cols-3 gap-1">
                          <div className="bg-emerald-50/70 border border-emerald-100 rounded px-1.5 py-1">
                            <div className="text-[8px] uppercase font-semibold tracking-wider text-emerald-700">Qty</div>
                            <div className="text-[11px] font-bold font-mono text-emerald-800">{stats.totalQty.toLocaleString("id-ID")}</div>
                          </div>
                          <div className="bg-indigo-50/70 border border-indigo-100 rounded px-1.5 py-1">
                            <div className="text-[8px] uppercase font-semibold tracking-wider text-indigo-700">Bin/TO</div>
                            <div className="text-[11px] font-bold font-mono text-indigo-800">{stats.avgBinPerTO}</div>
                          </div>
                          <div className="bg-amber-100/70 border border-amber-200 rounded px-1.5 py-1">
                            <div className="text-[8px] uppercase font-semibold tracking-wider text-amber-700">Qty/Bin</div>
                            <div className="text-[11px] font-bold font-mono text-amber-800">{stats.avgQtyPerBin.toFixed(1)}</div>
                          </div>
                        </div>
                      </div>
                    );
                  })()}
                </div>
              </div>
            );
          })()}

          {selectedWing === "cross" && (
            <div className="p-4 rounded-2xl border bg-rose-50/40 border-rose-200/60" data-testid="cross-wing-panel">
              <div className="text-xs font-semibold text-slate-700 mb-3 flex items-center">
                <Package className="w-3.5 h-3.5 mr-1.5 text-rose-600" />
                Cross Wing — pilih Batch Picking untuk membagi beban picker:
              </div>

              <div
                onClick={() => setBatchPickingOpen(!batchPickingOpen)}
                data-testid="subcard-batch-picking"
                className={`p-4 rounded-xl border cursor-pointer transition-all ${
                  batchPickingOpen
                    ? "bg-rose-50 border-rose-500 ring-2 ring-rose-400/30 shadow-md"
                    : "bg-white border-rose-200 hover:border-rose-400 hover:shadow-md"
                }`}
              >
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <div className={`p-2 rounded-lg ${batchPickingOpen ? "bg-rose-600 text-white" : "bg-rose-100 text-rose-700"}`}>
                      <Package className="w-4 h-4" />
                    </div>
                    <div>
                      <div className="text-sm font-bold text-slate-800">Batch Picking</div>
                      <div className="text-[11px] text-slate-500">Bagi {wingAssignments.cross.length} TO Cross Wing jadi batch, print pick ticket rata per picker.</div>
                    </div>
                  </div>
                  <ArrowRight className={`w-4 h-4 transition-transform ${batchPickingOpen ? "rotate-90 text-rose-600" : "text-rose-400"}`} />
                </div>
              </div>

              {batchPickingOpen && (
                <div className="mt-3 bg-white border border-rose-200/70 rounded-xl p-3.5" data-testid="batch-size-selector">
                  <div className="text-xs font-semibold text-slate-700 mb-2 flex items-center">
                    <Layers className="w-3.5 h-3.5 mr-1.5 text-rose-500" />
                    Pilih ukuran batch (jumlah TO per batch):
                  </div>
                  <div className="flex flex-wrap gap-2">
                    {[10, 15, 20, 25, 30].map(size => {
                      const active = batchSize === size;
                      return (
                        <button
                          key={size}
                          onClick={() => { setBatchSize(size); setCustomBatchSize(""); }}
                          data-testid={`batch-size-${size}`}
                          className={`text-xs font-semibold px-3 py-1.5 rounded-lg border transition-all ${
                            active ? "bg-rose-600 border-rose-600 text-white shadow-sm" : "bg-white border-rose-200 text-rose-800 hover:bg-rose-50"
                          }`}
                        >
                          {size} TO
                        </button>
                      );
                    })}
                    <div className="flex items-center gap-1">
                      <Input
                        type="number"
                        min={1}
                        max={200}
                        placeholder="Custom"
                        value={customBatchSize}
                        onChange={(e) => setCustomBatchSize(e.target.value)}
                        data-testid="batch-size-custom-input"
                        className="h-9 w-24 text-xs border-rose-200"
                      />
                      <Button
                        size="sm"
                        onClick={() => {
                          const n = parseInt(customBatchSize, 10);
                          if (!Number.isNaN(n) && n > 0) setBatchSize(n);
                          else toast.error("Custom batch size harus angka > 0");
                        }}
                        data-testid="batch-size-custom-apply"
                        className="h-9 text-xs bg-rose-600 hover:bg-rose-700"
                      >
                        Set
                      </Button>
                    </div>
                  </div>

                  {batchSize && batches.length > 0 && (
                    <>
                      <div className="text-xs text-slate-600 mt-3 mb-2">
                        Terbagi menjadi <span className="font-bold text-rose-700">{batches.length}</span> batch (@ {batchSize} TO). Pilih batch untuk print:
                      </div>
                      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-2" data-testid="batch-grid">
                        {batches.map(b => {
                          const active = selectedBatchNumber === b.batchNumber;
                          return (
                            <button
                              key={b.batchNumber}
                              onClick={() => setSelectedBatchNumber(active ? null : b.batchNumber)}
                              data-testid={`batch-card-${b.batchNumber}`}
                              className={`text-left rounded-lg border p-2.5 transition-all ${
                                active ? "bg-rose-600 border-rose-600 text-white shadow-sm" : "bg-white border-rose-200 hover:border-rose-400"
                              }`}
                            >
                              <div className={`text-[10px] uppercase font-semibold tracking-wider ${active ? "text-white/80" : "text-rose-500"}`}>Batch</div>
                              <div className={`text-lg font-bold font-mono ${active ? "text-white" : "text-rose-800"}`}>{b.batchNumber}</div>
                              <div className={`text-[10px] font-mono ${active ? "text-white/80" : "text-slate-500"}`}>
                                {b.tos.length} TO · Qty {b.stats.totalQty.toLocaleString("id-ID")}
                              </div>
                              <div className={`text-[9px] font-mono ${active ? "text-white/70" : "text-slate-400"}`}>
                                Kode {b.codedTOs[0]?.code} - {b.codedTOs[b.codedTOs.length - 1]?.code}
                              </div>
                            </button>
                          );
                        })}
                      </div>
                    </>
                  )}
                </div>
              )}
            </div>
          )}

          {selectedSubCard && selectedSubCard.type === "mixed" && (() => {
            const wingBucket = wingAssignments[selectedSubCard.wing];
            const mixed = wingBucket.mixed;
            const bucketAgg = (predicate) => mixed.reduce((acc, to) => {
              const info = toRowInfo.get(to);
              if (!info) return acc;
              const match = predicate === null ? true : predicate(info.numericRows.size);
              if (match) {
                acc.count += 1;
                acc.qty += info.qty;
              }
              return acc;
            }, { count: 0, qty: 0 });
            const all = bucketAgg(null);
            const wingShort = WING_STRUCTURE[selectedSubCard.wing].shortLabel;
            const buckets = [
              { key: "all", label: `Semua Mixed ${wingShort}`, ...all },
              { key: "2", label: "2 Row", ...bucketAgg(n => n === 2) },
              { key: "3", label: "3 Row", ...bucketAgg(n => n === 3) },
              { key: "4", label: "4 Row", ...bucketAgg(n => n === 4) },
              { key: "5", label: "5 Row", ...bucketAgg(n => n === 5) },
              { key: ">5", label: "> 5 Row", ...bucketAgg(n => n > 5) },
            ];
            return (
              <div
                className="bg-amber-50/60 border border-amber-200 rounded-2xl px-4 py-3"
                data-testid="mixed-subfilter-bar"
              >
                <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 mb-2.5">
                  <div className="flex items-center text-xs font-semibold text-amber-900">
                    <Filter className="w-3.5 h-3.5 mr-1.5 text-amber-600" />
                    Sub-filter Mixed {wingShort} — kelompokkan berdasarkan jumlah Row per TO:
                  </div>
                  <span className="text-[11px] text-amber-700 font-medium">
                    Aktif: <span className="font-bold">{buckets.find(b => b.key === mixedSubFilter)?.label}</span>
                  </span>
                </div>
                <div className="flex flex-wrap gap-2">
                  {buckets.map(b => {
                    const active = mixedSubFilter === b.key;
                    return (
                      <button
                        key={b.key}
                        type="button"
                        onClick={() => setMixedSubFilter(b.key)}
                        data-testid={`mixed-subfilter-${b.key === ">5" ? "gt5" : b.key}`}
                        className={`text-left rounded-xl border transition-all px-3 py-2 min-w-[112px] ${
                          active
                            ? "bg-amber-600 border-amber-600 text-white shadow-sm"
                            : "bg-white border-amber-200 text-amber-900 hover:bg-amber-100 hover:border-amber-400"
                        }`}
                      >
                        <div className={`text-xs font-bold ${active ? "text-white" : "text-amber-900"}`}>
                          {b.label}
                        </div>
                        <div className={`mt-1 flex items-center gap-1.5 text-[10px] font-mono ${active ? "text-white/95" : "text-amber-800"}`}>
                          <span
                            className={`px-1.5 py-0.5 rounded ${active ? "bg-white/20" : "bg-amber-50 border border-amber-200"}`}
                            data-testid={`mixed-subfilter-${b.key === ">5" ? "gt5" : b.key}-count`}
                          >
                            {b.count.toLocaleString("id-ID")} TO
                          </span>
                          <span
                            className={`px-1.5 py-0.5 rounded ${active ? "bg-white/20" : "bg-emerald-50 border border-emerald-200 text-emerald-800"}`}
                            data-testid={`mixed-subfilter-${b.key === ">5" ? "gt5" : b.key}-qty`}
                          >
                            Qty {b.qty.toLocaleString("id-ID")}
                          </span>
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>
            );
          })()}
          </>
          )}
        </div>

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
                    if (!selectedWing && searchQuery) {
                      return `Hasil pencarian untuk "${searchQuery}"`;
                    }
                    if (!selectedWing) {
                      return "Menampilkan semua Transfer Order Number unik (Pilih Wing di atas untuk memfilter)";
                    }
                    if (selectedWing === "cross") {
                      if (selectedBatch) {
                        return `Menampilkan TO Batch ${selectedBatch.batchNumber} (${selectedBatch.tos.length} TO · Kode ${selectedBatch.codedTOs[0]?.code}-${selectedBatch.codedTOs[selectedBatch.codedTOs.length-1]?.code}) — siap di-print sebagai Pick Ticket A4`;
                      }
                      return "Menampilkan TO Cross Wing — pilih Batch Picking di atas untuk membagi jadi batch, atau print pick ticket per batch";
                    }
                    const wingLabel = WING_STRUCTURE[selectedWing].label;
                    if (!selectedSubCard) {
                      return `Menampilkan SEMUA TO di ${wingLabel} (pilih sub-card untuk fokus ke rentang tertentu)`;
                    }
                    if (selectedSubCard.type === "range") {
                      return `Menampilkan TO ${wingLabel} yang SEMUA item Row-nya ada di rentang ${selectedSubCard.key}`;
                    }
                    if (selectedSubCard.type === "mixed") {
                      let bucketLabel;
                      if (mixedSubFilter === "all") {
                        bucketLabel = "semua";
                      } else if (mixedSubFilter === ">5") {
                        bucketLabel = "lebih dari 5 Row";
                      } else {
                        bucketLabel = `tepat ${mixedSubFilter} Row`;
                      }
                      return `Menampilkan TO Mixed ${WING_STRUCTURE[selectedSubCard.wing].shortLabel} (${bucketLabel}) — TO yang item Row-nya nyebar di lebih dari 1 sub-range wing`;
                    }
                    return "";
                  })()}
                </CardDescription>
              </div>

              <div className="flex items-center space-x-3">
                {selectedWing === "cross" && selectedBatch && (
                  <>
                    <Button
                      onClick={handleRequestPrint}
                      data-testid="print-pick-ticket-btn"
                      className="bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold shadow-xs flex items-center"
                    >
                      <Printer className="w-3.5 h-3.5 mr-1.5" />
                      Print Pick Ticket Batch {selectedBatch.batchNumber}
                    </Button>
                    
                    <Button
                      onClick={handlePrintChecklist}
                      data-testid="print-checklist-btn"
                      className="bg-[#1e3a8a] hover:bg-blue-900 text-white text-xs font-semibold shadow-xs flex items-center"
                    >
                      <CheckSquare className="w-3.5 h-3.5 mr-1.5" />
                      Print Form TO Checklist
                    </Button>
                  </>
                )}
                <Button
                  onClick={handleCopyToSAP}
                  data-testid="copy-to-sap-btn"
                  className={`bg-teal-600 hover:bg-teal-700 text-white text-xs font-semibold shadow-xs flex items-center transition-all ${copyPulse ? "ring-4 ring-emerald-300/60 scale-105 bg-emerald-600 hover:bg-emerald-600" : ""}`}
                >
                  {copyPulse ? (
                    <>
                      <Check className="w-3.5 h-3.5 mr-1.5" />
                      Tersalin!
                    </>
                  ) : (
                    <>
                      <Copy className="w-3.5 h-3.5 mr-1.5" />
                      Copy TO Unik ke SAP
                    </>
                  )}
                </Button>
              </div>
            </div>
          </CardHeader>
          <CardContent className="p-0">
            <div className="overflow-x-auto max-h-[500px]">
              <table className="w-full text-left border-collapse text-xs" data-testid="unique-to-table">
                <thead className="bg-slate-100/80 sticky top-0 z-10 text-slate-700 font-semibold border-b border-slate-200">
                  <tr>
                    <th className="py-3 px-4 w-16">No</th>
                    <th className="py-3 px-4 text-teal-800 font-bold bg-teal-50">Transfer Order Number</th>
                    {isFiltered && (
                      <>
                        <th className="py-3 px-4 text-indigo-800 font-bold bg-indigo-50 w-28">Total Bin</th>
                        <th className="py-3 px-4 text-rose-800 font-bold bg-rose-50 w-28">Total Article</th>
                        <th className="py-3 px-4 text-emerald-800 font-bold bg-emerald-50 w-28">Total Qty</th>
                        <th className="py-3 px-4 text-amber-800 font-bold bg-amber-50">Row Spread</th>
                      </>
                    )}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {uniqueTransferOrders.length > 0 ? (
                    uniqueTransferOrders.map((to, idx) => {
                      const info = toRowInfo.get(to);
                      const rows = info ? Array.from(info.numericRows).sort((a, b) => a - b) : [];
                      const totalBin = info ? info.bins.size : 0;
                      const totalArticle = info ? info.articles.size : 0;
                      const totalQty = info ? info.qty : 0;
                      return (
                        <tr key={to} data-testid={`to-row-${idx}`} className="hover:bg-slate-50/80 transition-colors">
                          <td className="py-2.5 px-4 text-slate-400 font-mono">{idx + 1}</td>
                          <td className="py-2.5 px-4 font-bold font-mono text-teal-800 text-sm">{to}</td>
                          {isFiltered && (
                            <>
                              <td className="py-2.5 px-4 bg-indigo-50/40 font-mono font-semibold text-indigo-800" data-testid={`total-bin-${idx}`}>{totalBin.toLocaleString("id-ID")}</td>
                              <td className="py-2.5 px-4 bg-rose-50/40 font-mono font-semibold text-rose-800" data-testid={`total-article-${idx}`}>{totalArticle.toLocaleString("id-ID")}</td>
                              <td className="py-2.5 px-4 bg-emerald-50/40 font-mono font-semibold text-emerald-800" data-testid={`total-qty-${idx}`}>{totalQty.toLocaleString("id-ID")}</td>
                              <td className="py-2.5 px-4 bg-amber-50/40" data-testid={`row-spread-${idx}`}>
                                <div className="flex flex-wrap gap-1">
                                  {rows.map(r => (
                                    <span key={r} className="text-[11px] font-mono bg-amber-100 text-amber-800 px-1.5 py-0.5 rounded border border-amber-200">
                                      {String(r).padStart(2, "0")}
                                    </span>
                                  ))}
                                </div>
                              </td>
                            </>
                          )}
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan={isFiltered ? "6" : "2"} className="py-12 text-center text-slate-400">
                        Tidak ada Transfer Order Number yang sesuai dengan filter saat ini.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

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

      {printModalOpen && selectedBatch && printWarning && (
        <div
          className="fixed inset-0 z-40 flex items-center justify-center bg-slate-900/40 backdrop-blur-sm no-print"
          data-testid="print-modal"
        >
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-md mx-4 overflow-hidden">
            <div className="flex items-start justify-between px-5 py-4 border-b border-slate-200 bg-amber-50">
              <div>
                <div className="text-sm font-bold text-slate-800 flex items-center">
                  <AlertTriangle className="w-4 h-4 mr-1.5 text-amber-600" />
                  Peringatan Re-Print · Batch {selectedBatch.batchNumber}
                </div>
                <div className="text-[11px] text-slate-500 font-mono mt-0.5">
                  {formatBatchCode(selectedBatch.batchNumber, selectedBatch.randomSuffix)} · {selectedBatch.tos.length} TO
                </div>
              </div>
              <button
                onClick={() => { setPrintModalOpen(false); setPrintWarning(null); }}
                className="text-slate-400 hover:text-slate-700"
                data-testid="print-modal-close"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5 space-y-3">
              <div className="rounded-lg border border-amber-300 bg-amber-50 p-3" data-testid="print-warning">
                <div className="flex items-start gap-2">
                  <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                  <div className="text-xs">
                    <div className="font-bold text-amber-900 mb-1">
                      {printWarning.previousPrints.length} TO sudah pernah di-print sebelumnya!
                    </div>
                    <div className="max-h-40 overflow-y-auto space-y-0.5 text-amber-800 font-mono text-[11px]">
                      {printWarning.previousPrints.slice(0, 30).map(p => (
                        <div key={p.to_number}>
                          <span className="font-bold">{p.to_number}</span> · {p.batch_code} · {new Date(p.printed_at).toLocaleString("id-ID")}
                        </div>
                      ))}
                      {printWarning.previousPrints.length > 30 && (
                        <div className="text-[10px] text-amber-600 italic">
                          ... dan {printWarning.previousPrints.length - 30} lainnya
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 px-5 py-3 border-t border-slate-200 bg-slate-50">
              <Button
                variant="outline"
                size="sm"
                onClick={() => { setPrintModalOpen(false); setPrintWarning(null); }}
                data-testid="print-modal-cancel"
                className="text-xs"
              >
                Batal
              </Button>
              <Button
                size="sm"
                onClick={() => { setPrintModalOpen(false); proceedWithPrint(); }}
                data-testid="print-force-btn"
                className="text-xs bg-amber-600 hover:bg-amber-700"
              >
                <Printer className="w-3.5 h-3.5 mr-1.5" />
                Tetap Print
              </Button>
            </div>
          </div>
        </div>
      )}

      {selectedBatch && printType === 'ptf' && (
        <div className="print-only pick-ticket-print" data-testid="pick-ticket-print" aria-hidden="true">
          {(() => {
            const rowPages = buildRowPages(selectedBatch);
            const batchCode = formatBatchCode(selectedBatch.batchNumber, selectedBatch.randomSuffix);
            const printedAt = new Date().toLocaleString("id-ID", { dateStyle: "short", timeStyle: "medium" });
            const codeMap = new Map(selectedBatch.codedTOs.map(c => [c.to, c.code]));

            return rowPages.map((page, pIdx) => {
              const rowLabel = (() => {
                const n = parseInt(page.rowKey, 10);
                return Number.isNaN(n) ? String(page.rowKey) : String(n).padStart(2, "0");
              })();

              return (
                <div key={`${page.rowKey}-${page.rowPageNumber}`} className="pick-page">
                  {page.isFirstPage && (
                    <>
                      <div className="pt-header">
                        <div className="pt-title">
                          <div className="pt-h1">TRANSFER ORDER PICKING</div>
                          <div className="pt-batch-badge">
                            <span className="pt-batch-label">Batch</span>
                            <span className="pt-batch-num">{selectedBatch.batchNumber}</span>
                            <span className="pt-batch-meta">· {selectedBatch.tos.length} TO in batch</span>
                          </div>
                        </div>

                        <div className="pt-row-badge">
                          <div className="pt-row-label">ROW</div>
                          <div className="pt-row-value">{rowLabel}</div>
                        </div>

                        <div className="pt-code-box">
                          {reprintCount > 0 && (
                            <div
                              style={{
                                fontSize: '22px',
                                fontWeight: '900',
                                border: '3px solid black',
                                padding: '4px 10px',
                                marginBottom: '8px',
                                textAlign: 'center',
                                letterSpacing: '1px',
                                backgroundColor: '#ffffff'
                              }}
                            >
                              RE-PRINT {reprintCount}
                            </div>
                          )}

                          <Barcode
                            value={batchCode}
                            width={1.2}
                            height={35}
                            displayValue={false}
                            margin={0}
                          />

                          <div className="pt-code-num" style={{ marginTop: '4px' }}>
                            {batchCode}
                          </div>

                          <div className="pt-code-meta">
                            Printed: {printedAt}
                          </div>
                        </div>
                      </div>

                      <div className="pt-summary">
                        <div className="pt-summary-item">
                          <span className="pt-sm-label">Total TO</span>
                          <span className="pt-sm-value">{page.stats.totalTO}</span>
                        </div>

                        <div className="pt-summary-item">
                          <span className="pt-sm-label">Total Article</span>
                          <span className="pt-sm-value">{page.stats.totalArticle}</span>
                        </div>

                        <div className="pt-summary-item">
                          <span className="pt-sm-label">Total Qty</span>
                          <span className="pt-sm-value">{page.stats.totalQty.toLocaleString("id-ID")}</span>
                        </div>
                      </div>
                    </>
                  )}

                  <table className="pt-table">
  <thead>
    <tr>
      <th style={{ width: "36px" }}>No. Urut</th>
      <th style={{ width: "78px" }}>Source Bin</th>
      <th style={{ width: "150px" }}>Article</th>
      <th>Article Description</th>
      <th style={{ width: "42px" }}>Qty</th>
      <th style={{ width: "37px" }}>UoM</th>
      <th style={{ width: "48px" }}>Kode</th>
      <th style={{ width: "47px" }}>TO Line</th>
      <th style={{ width: "78px" }}>TO Number</th>
      <th style={{ width: "78px" }}>Dest. Bin</th>
    </tr>
  </thead>
  <tbody>
    {page.items.map((it, idx) => (
      <tr key={idx}>
        <td>{String(idx + 1).padStart(4, "0")}</td>
        <td className="pt-nowrap">{it["Source Storage Bin"] || ""}</td>
        <td className="pt-article">{it["Article"] || ""}</td>
        <td className="pt-desc">{it["Article Description"] || ""}</td>
        <td>{it["Source target qty"] || ""}</td>
        <td>EA</td>
        <td style={{ fontWeight: 'bold', color: '#000000', fontSize: '14px' }}>
          {codeMap.get(it["Transfer Order Number"]) || ""}
        </td>
        <td>{it["Transfer order item"] || ""}</td>
        <td className="pt-nowrap">{it["Transfer Order Number"]}</td>
        <td className="pt-nowrap">{it["Dest.Storage Bin"] || ""}</td>
      </tr>
    ))}
  </tbody>
</table>

                  <div className="pt-footer">
                    <div className="pt-footer-row">
                      <div>
                        Total Items in this Row: <strong>{page.totalRowItems}</strong> · Row <strong>{rowLabel}</strong>
                      </div>
                      <div className="pt-page-num">
                        Page {page.rowPageNumber} of {page.rowTotalPages}
                      </div>
                    </div>

                    <div className="pt-signature">
                      <div>Picker: ____________________</div>
                      <div>Checker: ____________________</div>
                    </div>
                  </div>
                </div>
              );
            });
          })()}
        </div>
      )}

      {selectedBatch && printType === 'checklist' && (
        <div className="print-only" aria-hidden="true">
          {selectedBatch.codedTOs.map((item, pageIndex) => {
             const info = toRowInfo.get(item.to);
             // Urutkan row dari rendah ke tinggi, tanpa pisah wing
             const rows = info ? Array.from(info.numericRows).sort((a,b)=>a-b) : [];
             
             // Pastikan info.totalQty sudah dihitung di logic atas (saat build toRowInfo)
             // Jika belum, kasih fallback angka 0
             const totalQty = info?.totalQty || 0; 
             
             const today = new Date();
             const dateStr = today.toLocaleDateString('id-ID', { day: '2-digit', month: '2-digit', year: 'numeric' });
             const timeStr = today.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });

             return (
               <div key={item.to} className="checklist-a4-page flex flex-col" style={{ pageBreakAfter: 'always', breakAfter: 'page', minHeight: '100vh', backgroundColor: '#ffffff', padding: '40px' }}>
                 
                 {/* HEADER */}
                 <div className="flex border-b-[3px] border-slate-300 pb-6 mb-6">
                   <div className="w-1/2 flex items-center gap-6">
                     <div className="bg-[#1e3a8a] p-4 rounded-xl shadow-sm border-2 border-blue-900">
                       <Package className="w-16 h-16 text-white" />
                     </div>
                     <div className="flex flex-col">
                       <div className="text-[#1e3a8a] font-bold text-lg tracking-widest mb-1">BATCH PICKING NUMBER</div>
                       <div className="text-[#1e3a8a] font-black text-[50px] leading-none">BATCH {selectedBatch.batchNumber}</div>
                     </div>
                   </div>
                   
                   <div className="w-1/2 flex flex-col">
                     <div className="bg-[#1e3a8a] text-white text-center py-2 font-bold text-xl tracking-widest border-b-2 border-slate-800">
                       KODE TO (STATION)
                     </div>
                     <div className="text-center py-4 flex-grow flex items-center justify-center">
                       <div className="text-[#1e3a8a] font-black text-[100px] leading-none tracking-tight">
                         {item.code}
                       </div>
                     </div>
                   </div>
                 </div>

                 {/* TO NUMBER & QTY */}
                 <div className="flex border-b-[3px] border-slate-300 pb-8 mb-8">
                   <div className="w-3/5 flex items-start gap-6 border-r-2 border-slate-200 pr-6">
                     <div className="bg-[#16a34a] rounded-xl p-5 flex-shrink-0 shadow-sm border-2 border-green-700 mt-2">
                       <ClipboardList className="w-14 h-14 text-white" />
                     </div>
                     <div className="flex flex-col">
                       <div className="text-[#16a34a] font-bold text-xl mb-1 tracking-wide">TO / ORDER</div>
                       <div className="text-[#16a34a] font-black text-[50px] leading-none mb-2">{item.to}</div>
                       <div className="text-slate-700 font-bold text-base mb-4">Masukkan ke carton sesuai TO ini</div>
                       
                       {/* DIGANTI JADI BARCODE BIASA (CODE 128) */}
                       <div className="mt-2">
                         <Barcode 
                           value={String(item.to)} 
                           width={2} 
                           height={50} 
                           displayValue={false} 
                           margin={0} 
                         />
                       </div>
                     </div>
                   </div>
                   
                   <div className="w-2/5 flex flex-col items-center justify-center pl-6">
                     <div className="text-slate-800 font-bold text-3xl mb-4 tracking-wider">QTY TO</div>
                     <div className="text-black font-black text-[80px] leading-none">
                       {totalQty.toLocaleString("id-ID")}
                     </div>
                   </div>
                 </div>

                 {/* CHECKLIST ROW */}
                 <div className="flex-grow flex flex-col">
                   <div className="bg-[#1e3a8a] text-white text-center py-3 font-bold text-xl tracking-wider mb-6 shadow-sm">
                     CHECKLIST ROW – CENTANG JIKA SUDAH SELESAI
                   </div>
                   <div className="bg-white p-8 rounded-xl border-2 border-slate-300 shadow-sm flex-grow">
                     <div className="flex flex-wrap gap-x-12 gap-y-10 justify-center">
                       {rows.length > 0 ? rows.map(r => (
                         <div key={r} className="flex flex-col items-center gap-4">
                           <div className="text-[#1e3a8a] font-bold text-2xl">ROW {r}</div>
                           <div className="w-16 h-16 border-[4px] border-slate-800 rounded-lg bg-white shadow-inner"></div>
                         </div>
                       )) : (
                         <div className="text-slate-400 italic font-semibold text-xl">Tidak ada data picking untuk TO ini</div>
                       )}
                     </div>
                   </div>
                 </div>

                 {/* FOOTER */}
                 <div className="mt-8 border-t-[3px] border-slate-300 pt-6 flex items-center justify-end gap-12">
                   <div className="flex items-center gap-4">
                     <Calendar className="w-8 h-8 text-[#1e3a8a]" />
                     <div className="flex flex-col">
                       <div className="text-xs text-slate-500 font-bold tracking-wider">TANGGAL</div>
                       <div className="font-black text-lg text-slate-800">{dateStr}</div>
                     </div>
                   </div>

                   <div className="flex items-center gap-4">
                     <Clock className="w-8 h-8 text-[#1e3a8a]" />
                     <div className="flex flex-col">
                       <div className="text-xs text-slate-500 font-bold tracking-wider">WAKTU</div>
                       <div className="font-black text-lg text-slate-800">{timeStr}</div>
                     </div>
                   </div>

                   <div className="text-lg font-bold text-slate-500 border-l-[3px] border-slate-300 pl-12 py-2">
                     Page {pageIndex + 1} of {selectedBatch.codedTOs.length}
                   </div>
                 </div>
                 
               </div>
             );
          })}
        </div>
      )}

      {uploadStatus.active && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 backdrop-blur-sm"
          data-testid="upload-loading-overlay"
          role="status"
          aria-live="polite"
        >
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 px-8 py-7 w-full max-w-md mx-4">
            <div className="flex items-center space-x-4">
              <div className="relative flex items-center justify-center h-14 w-14 rounded-full bg-teal-50 border border-teal-200 shrink-0">
                <Loader2 className="w-7 h-7 text-teal-600 animate-spin" />
              </div>
              <div className="flex-1 min-w-0">
                <div className="text-sm font-bold text-slate-800 truncate">
                  {uploadStatus.phase === "reading" && "Membaca file..."}
                  {uploadStatus.phase === "parsing" && "Memparsing sheet..."}
                  {uploadStatus.phase === "generating" && "Generate kolom Row & Card..."}
                </div>
                <div className="text-xs text-slate-500 truncate mt-0.5" title={uploadStatus.fileName}>
                  {uploadStatus.fileName}
                </div>
                {uploadStatus.rowsParsed > 0 && (
                  <div className="text-[11px] text-teal-700 font-semibold mt-1">
                    {uploadStatus.rowsParsed.toLocaleString("id-ID")} baris terdeteksi
                  </div>
                )}
              </div>
            </div>
            
            <p className="mt-3 text-[11px] text-slate-400 text-center">
              Jangan tutup tab ini sampai selesai ya, bro.
            </p>
          </div>
        </div>
      )}
    </div>
  );
}