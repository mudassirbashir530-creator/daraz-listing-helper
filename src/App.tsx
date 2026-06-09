import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import {
  Upload,
  FileSpreadsheet,
  Play,
  Download,
  Code,
  Copy,
  Check,
  Search,
  ArrowLeft,
  HelpCircle,
  Chrome,
  Bolt,
  Sparkles,
  RefreshCw,
  FileText,
  CheckCircle,
  Grid,
  FileCode,
  ExternalLink
} from 'lucide-react';

// Include extension static files from our code module
import {
  manifestCode,
  popupHtmlCode,
  popupJsCode,
  sidebarHtmlCode,
  sidebarJsCode,
  contentJsCode,
  backgroundJsCode,
  stylesCssCode,
  readmeCode
} from './extensionCode';

// Interfaces for our product models
interface Variation {
  pack: string;
  actualPrice: string;
  discountedPrice: string;
}

interface Product {
  serial: string;
  title: string;
  size: string;
  color: string;
  basePrice: string;
  variations: Variation[];
}

// Procedural generator for 33 mock products matching user's warehouse database
const getPrebakedProducts = (): Product[] => {
  const titles = [
    "Premium Square Acrylic Wall Mirror Tiles",
    "Classic Hexagonal Wall Mirror Sticker Set",
    "Modernist Chevron Grid Mirror Decals",
    "LED Backlit Vanity Mirror Light Bar (Dimmable)",
    "Islamic Calligraphy Bismillah Frame Art",
    "Minimalist Floating Wooden Pane Shelf Pack",
    "Round Beveled Mirror Adhesive Panels",
    "Retro Golden Framed Oval Wall Sticker",
    "3D Mirror Butterfly Wall Decor Accent Set",
    "Geometric Acrylic Silent Quartz Wall Clock",
    "Creative Wave Border Frameless Wall Mirror",
    "Starry Night Glowing Ceiling Stickers Pack",
    "Luminous Moon Phases Wall Decal Set",
    "Dandelion Whisper Mirror Stickers Set",
    "Heart-Shaped Wall Collage Panels (DIY)",
    "Self-Adhesive Real Glass Square Pack",
    "Abstract Ripple Acrylic Mirror Panel Sheet",
    "Boho Chic Rattan-Look Wall Decals Pack",
    "Hanging Chain Makeup Mirror Panel Frame",
    "DIY Corner Accent Scroll Decorative Decals",
    "Herringbone Acrylic Kitchen Splash Panel",
    "Deluxe Sunburst Design Accent Wall Mirror",
    "Retro Baroque Gold Mirror Frame Stencil",
    "Art Deco Triangle Pattern Glass Tiles",
    "Diamond Lattice Mirror Stickers Decal",
    "Kids Height Dinosaur Growth Measurement Chart",
    "Forest Animals Nursery Sticker Accents",
    "Universe Galaxy Ceiling Stars Glowers Pack",
    "Modern Honeycomb Hexagonal Shelf Panels",
    "Nordic Style Hanging Mirror Leather Strap",
    "Elegant Damask Mirror Border Trim Set",
    "Mandala Round Accent Mirror Sticker Decor",
    "Sleek Rectangular Closet Mirror Tiles"
  ];

  const colors = ["Silver Chrome", "Royal Gold", "Matte Black", "Amber Glow", "Rose Gold", "Bronze Gloss"];
  const sizes = ["4x4 Inches", "6x6 Inches", "8x8 Inches", "12x12 Inches", "Medium (18\" Dia)", "Large (24\" Dia)"];

  return titles.map((title, i) => {
    const sNum = 4051 + i;
    const baseVal = Math.floor((Math.random() * 50) + 20) * 5; // PKR base cost e.g. 100-350
    const variationsCount = 4; // Consistent pack sizes (4, 8, 12, 16)
    const variations: Variation[] = [];
    
    for (let j = 1; j <= variationsCount; j++) {
      const multiplier = j === 1 ? 4 : (j === 2 ? 8 : (j === 3 ? 12 : 16));
      // Actual Price formula (spans Row merged items)
      const act = baseVal * multiplier * 2.5;
      const disc = Math.round(act * 0.75); // 25% discount
      variations.push({
        pack: `Pack of ${multiplier}`,
        actualPrice: String(act),
        discountedPrice: String(disc)
      });
    }

    return {
      serial: String(sNum),
      title,
      size: sizes[i % sizes.length],
      color: colors[i % colors.length],
      basePrice: String(baseVal),
      variations
    };
  });
};

export default function App() {
  const [activeTab, setActiveTab] = useState<'simulator' | 'excel' | 'code'>('simulator');
  
  // Products states
  const [products, setProducts] = useState<Product[]>([]);
  const [loadedFilename, setLoadedFilename] = useState<string>('');
  
  // Core Extension Simulator States (represents state in popup.js running in fake chrome frame)
  const [simView, setSimView] = useState<'upload' | 'list' | 'detail'>('upload');
  const [simSearchQuery, setSimSearchQuery] = useState('');
  const [simActiveProduct, setSimActiveProduct] = useState<Product | null>(null);
  const [simToast, setSimToast] = useState<string | null>(null);
  const [simLastCopiedValue, setSimLastCopiedValue] = useState<string>('');
  const [simMode, setSimMode] = useState<'popup' | 'sidebar'>('popup');

  // Simulated Web Browser page (Daraz.pk Seller Form) States
  const [focusedField, setFocusedField] = useState<string | null>(null);
  const [darazFormState, setDarazFormState] = useState({
    title: '',
    size: '',
    color: '',
    basePrice: '',
    dimL: '',
    dimW: '',
    dimH: '',
    weight: '',
    warranty: '',
    returnPolicy: '',
    variationPrices: {} as Record<string, { act: string, sale: string }>
  });

  // Code Tab state for visual reviews
  const [selectedCodeTab, setSelectedCodeTab] = useState('manifest');
  const [copyCodeSuccess, setCopyCodeSuccess] = useState(false);

  // Sparkle feedback highlights for mock form fills
  const [sparkleField, setSparkleField] = useState<string | null>(null);

  // References for locating custom elements
  const dropzoneRef = useRef<HTMLDivElement>(null);

  // Load the prebaked sample database on start so they don't see blank views
  useEffect(() => {
    const prebaked = getPrebakedProducts();
    setProducts(prebaked);
    setLoadedFilename('Default Prebaked Database (33 products)');
  }, []);

  // Sync products list in simulation with parent products
  const loadWorkspaceProductsIntoSimulator = () => {
    setSimView('list');
    setSimToast('Products database loaded!');
    setTimeout(() => setSimToast(null), 1800);
  };

  // Drag and drop parser for actual excel sheets inside our browser app
  const onFileDrop = (e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      parseUploadedExcel(e.dataTransfer.files[0]);
    }
  };

  const onFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      parseUploadedExcel(e.target.files[0]);
    }
  };

  const parseUploadedExcel = (file: File) => {
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        if (!evt.target?.result) return;
        const bArray = new Uint8Array(evt.target.result as ArrayBuffer);
        const workbook = XLSX.read(bArray, { type: 'array' });
        
        const ws = workbook.Sheets[workbook.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as any[][];
        
        const productsList: Product[] = [];
        let current: Product | null = null;

        for (let i = 1; i < raw.length; i++) {
          const row = raw[i];
          if (!row || row.length === 0) continue;

          const serial = row[0];
          const title = row[1];
          const size = row[2];
          const color = row[3];
          const pack = row[4];
          const basePrice = row[5];
          const actualPrice = row[6];
          const discountedPrice = row[7];

          // Merged cell logic detector matching popup.js
          if (serial !== null && serial !== undefined && String(serial).trim() !== '') {
            if (current) productsList.push(current);
            current = {
              serial: String(serial).trim(),
              title: title ? String(title).trim() : '—',
              size: size ? String(size).trim() : '—',
              color: color ? String(color).trim() : '—',
              basePrice: basePrice ? String(basePrice).trim() : '—',
              variations: []
            };
          }

          if (current && pack !== null && pack !== undefined && String(pack).trim() !== '') {
            current.variations.push({
              pack: String(pack).trim(),
              actualPrice: actualPrice !== null && actualPrice !== undefined ? String(actualPrice).trim() : '—',
              discountedPrice: discountedPrice !== null && discountedPrice !== undefined ? String(discountedPrice).trim() : '0'
            });
          }
        }
        
        if (current) productsList.push(current);

        if (productsList.length === 0) {
          alert('No valid products parsed! Ensure Column A (Serial Number), B (Title), E (Pack Sizing), G (Actual Price) are correct.');
          return;
        }

        setProducts(productsList);
        setLoadedFilename(file.name);
        
        // Load into simulation state
        setSimView('list');
        setSimToast(`Successfully parsed ${productsList.length} products!`);
        setTimeout(() => setSimToast(null), 2500);

      } catch (err: any) {
        alert('Failed parsing spreadsheet: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Generate a premium specimen Excel workbook and trigger browser download
  const downloadSampleXLSX = () => {
    const data = [
      // Headers
      [
        "Product Serial Number",
        "Professional Title",
        "Size / Dimensions",
        "Color Options",
        "Variation Slot Name",
        "Base Unit Price (PKR)",
        "Actual Price (Daraz Cut)",
        "Discounted Price (Special Price)"
      ]
    ];

    const prebakedList = getPrebakedProducts();
    
    prebakedList.forEach((p, idx) => {
      p.variations.forEach((v, vIdx) => {
        if (vIdx === 0) {
          // Row 1 elements
          data.push([
            p.serial,
            p.title,
            p.size,
            p.color,
            v.pack,
            p.basePrice,
            v.actualPrice,
            v.discountedPrice
          ]);
        } else {
          // Row spans empty fillers
          data.push([
            null,
            null,
            null,
            null,
            v.pack,
            null,
            v.actualPrice,
            v.discountedPrice
          ]);
        }
      });
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "LIFEINNO Product List");
    XLSX.writeFile(wb, "LIFEINNO_Daraz_Products_Standard_Database.xlsx");
  };

  // Package entirely as full chrome extension and download
  const handleDownloadZipOfExtension = () => {
    const zip = new JSZip();
    
    zip.file("manifest.json", manifestCode);
    zip.file("popup.html", popupHtmlCode);
    zip.file("popup.js", popupJsCode);
    zip.file("sidebar.html", sidebarHtmlCode);
    zip.file("sidebar.js", sidebarJsCode);
    zip.file("content.js", contentJsCode);
    zip.file("background.js", backgroundJsCode);
    zip.file("styles.css", stylesCssCode);
    zip.file("README.md", readmeCode);
    
    // Tiny genuine valid orange png block base64 image (safeguarding loading asset warnings)
    const base64Png = "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAMAAAD8969fAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAADUExURfhdJKmH2K0AAAAGdFJOUVMAAAAAAIu627oAAAAkSURBVHja7cEBDAMAAAEg+Ve9pZcQAAAAAAAAAAAAAAAAAAAAAOBq694AH96TqZ0AAAAASUVORK5CYII=";
    zip.file("icon.png", base64Png, { base64: true });
    
    zip.generateAsync({ type: "blob" }).then((content) => {
      const link = document.createElement("a");
      link.href = URL.createObjectURL(content);
      link.download = "lifeinno_daraz_helper_chrome_extension.zip";
      link.click();
    });
  };

  // Helper inside mock Simulator detail view representing the Copy buttons
  const triggerMockCopy = (value: string, label: string) => {
    navigator.clipboard.writeText(value).catch(() => {});
    setSimLastCopiedValue(value);
    setSimToast(`Copied ${label}: "${value.length > 25 ? value.slice(0, 25) + '...' : value}"`);
    setTimeout(() => setSimToast(null), 2000);
  };

  // Helper inside mock Simulator detail view representing the ⚡ Fill active fields message
  const triggerMockAutofillField = (value: string, fieldKey: string) => {
    if (!value) return;
    setSimLastCopiedValue(value);

    // Active field determines where to write
    let targetField = focusedField;
    if (!targetField) {
      if (['title', 'size', 'color', 'basePrice', 'warranty', 'returnPolicy', 'dimL', 'dimW', 'dimH', 'weight'].includes(fieldKey)) {
        targetField = fieldKey;
      } else {
        // Fallback info notification
        setSimToast("⚠️ Focus on any text field in Daraz form first!");
        setTimeout(() => setSimToast(null), 2500);
        return;
      }
    }

    setDarazFormState(prev => ({
      ...prev,
      [targetField as string]: value
    }));

    setSparkleField(targetField);
    setTimeout(() => setSparkleField(null), 1200);

    setSimToast(`⚡ Injected into "${targetField}"!`);
    setTimeout(() => setSimToast(null), 1800);
  };

  // Calculated auto weight dimensions inside the simulator matching Feature 3
  const getSimulatedDimensions = (product: Product) => {
    const smallestPack = product.variations[0];
    const priceVal = smallestPack ? parseFloat(smallestPack.discountedPrice) : 0;
    
    if (priceVal <= 2000) {
      return { length: '20', width: '10', height: '5', weight: '300 grams' };
    } else {
      return { length: '30', width: '20', height: '10', weight: '700 grams' };
    }
  };

  // Copy code content string triggers in Visual Code viewer
  const getCodeTabString = () => {
    switch (selectedCodeTab) {
      case 'manifest': return manifestCode;
      case 'popup_html': return popupHtmlCode;
      case 'popup_js': return popupJsCode;
      case 'sidebar_html': return sidebarHtmlCode;
      case 'sidebar_js': return sidebarJsCode;
      case 'content': return contentJsCode;
      case 'background': return backgroundJsCode;
      case 'styles': return stylesCssCode;
      case 'readme': return readmeCode;
      default: return '';
    }
  };

  const getCodeTabFilename = () => {
    switch (selectedCodeTab) {
      case 'manifest': return 'manifest.json';
      case 'popup_html': return 'popup.html';
      case 'popup_js': return 'popup.js';
      case 'sidebar_html': return 'sidebar.html';
      case 'sidebar_js': return 'sidebar.js';
      case 'content': return 'content.js';
      case 'background': return 'background.js';
      case 'styles': return 'styles.css';
      case 'readme': return 'README.md';
      default: return '';
    }
  };

  const handleCopyCodeText = () => {
    const code = getCodeTabString();
    navigator.clipboard.writeText(code).then(() => {
      setCopyCodeSuccess(true);
      setTimeout(() => setCopyCodeSuccess(false), 2000);
    });
  };

  // Filtered products list for simulator search
  const filteredSimProducts = products.filter(p => {
    if (!simSearchQuery) return true;
    const q = simSearchQuery.toLowerCase();
    return (
      p.serial.toLowerCase().includes(q) ||
      p.title.toLowerCase().includes(q) ||
      p.size.toLowerCase().includes(q) ||
      p.color.toLowerCase().includes(q)
    );
  });

  // Calculate some stats
  const totalVariationsCount = products.reduce((acc, p) => acc + p.variations.length, 0);

  return (
    <div className="min-h-screen bg-[#F0F2F5] text-gray-800 flex flex-col font-sans">
      
      {/* HEADER HERO ELEMENT */}
      <header className="bg-white border-b border-gray-200 shadow-sm px-6 py-3.5 flex flex-col md:flex-row justify-between items-center gap-4">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-[#E85D24] text-white font-black text-2xl flex items-center justify-center rounded-xl shadow-md cursor-pointer transition-transform hover:scale-105">
            L
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="text-xl font-extrabold tracking-tight text-gray-900 font-display">
                LIFEINNO <span className="text-[#E85D24]">Daraz Helper</span>
              </h1>
              <span className="bg-[#E85D24]/10 text-[#E85D24] text-xs font-bold px-2 py-0.5 rounded border border-[#E85D24]/30 font-mono">
                CHROME EXTENSION DESK
              </span>
            </div>
            <p className="text-xs text-gray-500">
              Interactive extension testing environment, real-time workbook parser and packing studio.
            </p>
          </div>
        </div>

        {/* Global tab Switcher */}
        <div className="bg-gray-100 p-1 rounded-xl border border-gray-200 flex items-center gap-1">
          <button
            id="tab-btn-simulator"
            onClick={() => setActiveTab('simulator')}
            className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer ${
              activeTab === 'simulator'
                ? 'bg-[#E85D24] text-white border border-[#E85D24]/10 shadow-sm'
                : 'text-gray-600 hover:text-gray-900'
            }`}
          >
            <Play className="w-3.5 h-3.5" />
            Interactive Simulator
          </button>
          <button
            id="tab-btn-excel"
            onClick={() => setActiveTab('excel')}
            className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer ${
              activeTab === 'excel'
                ? 'bg-[#E85D24] text-white border border-[#E85D24]/10 shadow-sm'
                : 'text-gray-600 hover:text-gray-900'
            }`}
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            Excel Manager
          </button>
          <button
            id="tab-btn-code"
            onClick={() => setActiveTab('code')}
            className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer ${
              activeTab === 'code'
                ? 'bg-[#E85D24] text-white border border-[#E85D24]/10 shadow-sm'
                : 'text-gray-600 hover:text-gray-900'
            }`}
          >
            <Code className="w-3.5 h-3.5" />
            Code Exporter (.ZIP)
          </button>
        </div>
      </header>

      {/* VIEWPORT CONTROLLER */}
      <main className="flex-1 p-4 lg:p-6 overflow-hidden">
        
        {/* ======================= TAB 1: INTEGRATED SIMULATION ARENA ======================= */}
        {activeTab === 'simulator' && (
          <div className="h-full flex flex-col gap-6">
            
            {/* Top brief */}
            <div className="bg-white border border-gray-200 shadow-sm rounded-xl p-5 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
              <div>
                <h3 className="font-extrabold text-sm text-gray-900 flex items-center gap-2 font-display">
                  <Sparkles className="w-4 h-4 text-[#E85D24]" />
                  Verify Before Chrome Installation
                </h3>
                <p className="text-xs text-gray-500 max-w-xl mt-1 leading-relaxed">
                  Test the exact extension behaviors inside this mock browser! Click fields inside the **Daraz Form** to focus them, then use the **Simulator Widget** to search products, copy fields, or trigger instant active field injection.
                </p>
              </div>

              {/* In-tab database summary */}
              <div className="flex items-center gap-3 bg-gray-50 px-3.5 py-2 rounded-lg border border-gray-200">
                <div className="text-right">
                  <div className="text-xs font-mono font-bold text-gray-800">
                    {products.length} Products | {totalVariationsCount} SKUs
                  </div>
                  <div className="text-[10px] text-gray-400 max-w-[170px] truncate">
                    Loaded: {loadedFilename}
                  </div>
                </div>
                <button
                  id="btn-trigger-reload"
                  onClick={() => {
                    setSimView('upload');
                    setSimActiveProduct(null);
                  }}
                  className="bg-[#E85D24] hover:bg-[#d14f1b] text-white p-1.5 rounded-lg transition-colors cursor-pointer shadow-sm"
                  title="Reload xlsx template"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                </button>
              </div>
            </div>

            {/* Split layout: Extension simulation along with mock website page */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-stretch">
              
              {/* INTERACTIVE EXTENSION PANEL */}
              <div className="lg:col-span-4 flex flex-col gap-3">
                
                {/* Panel state toggler (Popup popover vs side panel) */}
                <div className="flex justify-between items-center bg-white px-3.5 py-2.5 rounded-lg border border-gray-200 shadow-sm">
                  <span className="text-xs font-bold text-gray-700 flex items-center gap-1.5 font-mono">
                    <Chrome className="w-3.5 h-3.5 text-[#E85D24]" />
                    Simulated Device
                  </span>
                  <div className="flex bg-gray-100 p-0.5 rounded-md text-[10px]">
                    <button
                      id="opt-popup-mode"
                      onClick={() => {
                        setSimMode('popup');
                        setSimView('list');
                      }}
                      className={`px-2.5 py-1 rounded cursor-pointer transition-all ${
                        simMode === 'popup' ? 'bg-white text-gray-900 shadow-sm font-bold' : 'text-gray-500 hover:text-gray-800'
                      }`}
                    >
                      Popup Over
                    </button>
                    <button
                      id="opt-sidebar-mode"
                      onClick={() => {
                        setSimMode('sidebar');
                        setSimView('list');
                      }}
                      className={`px-2.5 py-1 rounded cursor-pointer transition-all ${
                        simMode === 'sidebar' ? 'bg-white text-gray-900 shadow-sm font-bold' : 'text-gray-500 hover:text-gray-800'
                      }`}
                    >
                      Side Panel
                    </button>
                  </div>
                </div>

                {/* THE VIRTUAL CHROMED CONTAINER */}
                <div 
                  className={`bg-white border border-gray-200 rounded-2xl flex flex-col shadow-2xl relative transition-all mx-auto w-full max-w-[380px] overflow-hidden ${
                    simMode === 'popup' ? 'h-[520px]' : 'h-[620px]'
                  }`}
                  style={{ color: '#1A1A1A' }}
                >
                  
                  {/* Virtual Chrome Frame Header */}
                  <div className="bg-[#E85D24] text-white px-4.5 py-3 flex items-center justify-between border-b border-black/10 select-none shadow-sm">
                    <div className="flex items-center gap-2">
                      <div className="w-6 h-6 bg-white/20 text-white font-black text-xs flex items-center justify-center rounded-lg">L</div>
                      <span className="font-extrabold text-xs tracking-wider font-display">
                        {simMode === 'popup' ? 'LIFEINNO Daraz Helper' : 'LIFEINNO Side Helper'}
                      </span>
                    </div>
                    {simMode === 'popup' && (
                      <button
                        id="sim-action-openside"
                        onClick={() => {
                          setSimMode('sidebar');
                          setSimToast('Switched to side panel mode!');
                          setTimeout(() => setSimToast(null), 1500);
                        }}
                        className="bg-white/20 hover:bg-white/30 text-[9px] uppercase font-extrabold px-2 py-1 rounded-md border border-white/25 transition-all cursor-pointer"
                      >
                        Side Panel ➔
                      </button>
                    )}
                  </div>

                  {/* VIRTUAL POPUP / SIDEBAR DISPLAY CONTROLLER */}
                  <div className="flex-1 flex flex-col overflow-y-auto bg-white p-4 custom-scrollbar relative">
                    
                    {/* SIM VIEW 1: MANUAL FILE INGESTION */}
                    {simView === 'upload' && (
                      <div className="flex flex-col gap-4 text-center py-6">
                        <div className="w-12 h-12 bg-[#FFF5F2] border border-[#E85D24]/20 text-[#E85D24] text-xl font-black flex items-center justify-center rounded-2xl mx-auto shadow-sm">
                          L
                        </div>
                        <div>
                          <h4 className="text-xs font-extra-bold text-gray-900 font-display uppercase tracking-wider">Daraz Seller Assistant</h4>
                          <p className="text-[10px] text-gray-500 mt-1.5 px-4 leading-relaxed">
                            Upload your product database Excel sheet to enable active copy lists on active listing pages.
                          </p>
                        </div>

                        {/* Interactive Drag Drop Simulator */}
                        <div 
                          className="border-2 border-dashed border-gray-200 hover:border-[#E85D24] bg-gray-50 hover:bg-[#FFF5F2]/50 rounded-xl p-5 cursor-pointer transition-all duration-150 shadow-sm"
                          onClick={() => document.getElementById('sim-file-input')?.click()}
                        >
                          <FileSpreadsheet className="w-8 h-8 text-gray-400 mx-auto mb-2" />
                          <span className="text-xs font-bold text-gray-800 block">Choose Excel Template</span>
                          <span className="text-[9px] text-gray-400 mt-1 block">Compatible standard columns A-H</span>
                          <input 
                             type="file" 
                             id="sim-file-input" 
                             className="hidden" 
                             accept=".xlsx, .xls"
                             onChange={onFileInputChange} 
                          />
                        </div>

                        <div className="text-[10px] text-gray-600 bg-gray-50 p-3 rounded-xl border border-gray-150">
                          Or bypass using preloaded database:
                          <button
                            id="btn-import-mock"
                            onClick={loadWorkspaceProductsIntoSimulator}
                            className="w-full bg-[#E85D24] hover:bg-[#d14f1b] text-white mt-2 py-2 rounded-lg text-[10px] font-bold cursor-pointer transition-colors shadow-sm"
                          >
                            🚀 Load Prebaked 33 Products
                          </button>
                        </div>
                      </div>
                    )}

                    {/* SIM VIEW 2: PRODUCT CARD VIEWLIST */}
                    {simView === 'list' && (
                      <div className="flex flex-col h-full">
                        {/* Search Container */}
                        <div className="relative flex items-center mb-2">
                          <Search className="w-3.5 h-3.5 text-gray-400 absolute left-2.5 pointer-events-none" />
                          <input
                            type="text"
                            placeholder="Search title, serial, size..."
                            value={simSearchQuery}
                            onChange={(e) => setSimSearchQuery(e.target.value)}
                            className="w-full pl-8 pr-3 py-1.5 border border-gray-200 rounded-md text-xs bg-gray-50 focus:bg-white focus:ring-1 focus:ring-[#E85D24] outline-none"
                          />
                        </div>

                        <div className="flex justify-between items-center mb-2 text-[10px]">
                          <span className="font-semibold text-[#E85D24]">
                            {filteredSimProducts.length} items parsed
                          </span>
                          <button
                            onClick={() => setSimView('upload')}
                            className="text-[#6B7280] hover:text-[#E85D24] underline cursor-pointer"
                          >
                            Change File
                          </button>
                        </div>

                        {/* List items */}
                        <div className="flex-1 flex flex-col gap-2 overflow-y-auto max-h-[460px] pr-1">
                          {filteredSimProducts.map((p, idx) => (
                            <div
                              key={p.serial}
                              onClick={() => {
                                setSimActiveProduct(p);
                                setSimView('detail');
                              }}
                              className="border border-gray-200 hover:border-[#E85D24] bg-white hover:bg-[#FFF5F2] hover:shadow-sm p-2 rounded-lg cursor-pointer transition-all"
                            >
                              <div className="flex justify-between items-center text-[10px] mb-1">
                                <span className="font-mono font-bold bg-gray-100 text-gray-600 px-1 rounded">#{p.serial}</span>
                                <span className="text-[#E85D24] font-bold">{p.variations.length} packs</span>
                              </div>
                              <h5 className="font-bold text-xs text-gray-900 line-clamp-2 leading-tight">{p.title}</h5>
                              <div className="text-[10px] text-gray-400 mt-1 flex gap-2">
                                <span>Size: {p.size}</span>
                                <span>|</span>
                                <span>Col: {p.color}</span>
                              </div>
                            </div>
                          ))}

                          {filteredSimProducts.length === 0 && (
                            <div className="text-center py-12 text-gray-400 text-xs">
                              No matching products found.
                            </div>
                          )}
                        </div>
                      </div>
                    )}

                    {/* SIM VIEW 3: RICH DETAILS WITH ATOM COPY ACTIONS */}
                    {simView === 'detail' && simActiveProduct && (
                      <div className="flex flex-col h-full text-gray-800">
                        
                        {/* Navigation back and header */}
                        <div className="flex justify-between items-center border-b border-gray-100 pb-2 mb-2">
                          <button
                            onClick={() => setSimView('list')}
                            className="text-xs font-bold text-[#E85D24] hover:underline flex items-center gap-1 cursor-pointer"
                          >
                            ← Products list
                          </button>
                          <span className="font-mono text-[10px] font-bold bg-gray-800 text-white px-2 py-0.5 rounded">
                            SN: {simActiveProduct.serial}
                          </span>
                        </div>

                        {/* Field list and Action triggers */}
                        <div className="flex flex-col gap-2.5 overflow-y-auto pr-1">
                             {/* Title element */}
                          <div className="border border-gray-200 rounded-xl p-3 bg-gray-50/50 shadow-sm">
                            <label className="text-[8.5px] uppercase tracking-wider font-extrabold text-[#E85D24] block mb-1">Product Title</label>
                            <div className="flex justify-between items-start gap-3">
                              <span className="text-xs text-gray-800 font-extrabold flex-1 leading-tight">{simActiveProduct.title}</span>
                              <div className="flex gap-1 shrink-0">
                                <button
                                  onClick={() => triggerMockCopy(simActiveProduct.title, 'Title')}
                                  className="w-6.5 h-6.5 bg-white border border-gray-200 hover:bg-gray-100 rounded-lg flex items-center justify-center text-xs shadow-sm cursor-pointer"
                                  title="Copy field text"
                                >
                                  📋
                                </button>
                                <button
                                  onClick={() => triggerMockAutofillField(simActiveProduct.title, 'title')}
                                  className="w-6.5 h-6.5 bg-gray-150 hover:bg-[#E85D24] border border-gray-200/50 hover:border-[#E85D24] text-gray-600 hover:text-white rounded-lg flex items-center justify-center text-xs font-mono font-bold transition-all cursor-pointer shadow-sm"
                                  title="Fill active web field"
                                >
                                  ⚡
                                </button>
                              </div>
                            </div>
                          </div>

                          {/* Dual Row Size & Color */}
                          <div className="grid grid-cols-2 gap-2.5">
                            <div className="border border-gray-200 rounded-xl p-3 bg-gray-50/50 shadow-sm">
                              <label className="text-[8.5px] uppercase tracking-wider font-extrabold text-gray-400 block mb-1">Dimension Size</label>
                              <div className="flex justify-between items-center gap-1">
                                <span className="text-xs font-bold text-gray-800 truncate">{simActiveProduct.size}</span>
                                <div className="flex gap-1 shrink-0">
                                  <button
                                    onClick={() => triggerMockCopy(simActiveProduct.size, 'Size')}
                                    className="w-5.5 h-5.5 bg-white border border-gray-250 hover:bg-gray-100 rounded-md flex items-center justify-center text-[10px] cursor-pointer shadow-sm"
                                  >
                                    📋
                                  </button>
                                  <button
                                    onClick={() => triggerMockAutofillField(simActiveProduct.size, 'size')}
                                    className="w-5.5 h-5.5 bg-gray-150 hover:bg-[#E85D24] text-gray-600 hover:text-white rounded-md flex items-center justify-center text-[10px] cursor-pointer transition-colors"
                                  >
                                    ⚡
                                  </button>
                                </div>
                              </div>
                            </div>

                            <div className="border border-gray-200 rounded-xl p-3 bg-gray-50/50 shadow-sm">
                              <label className="text-[8.5px] uppercase tracking-wider font-extrabold text-gray-400 block mb-1">Color options</label>
                              <div className="flex justify-between items-center gap-1">
                                <span className="text-xs font-bold text-gray-800 truncate">{simActiveProduct.color}</span>
                                <div className="flex gap-1 shrink-0">
                                  <button
                                    onClick={() => triggerMockCopy(simActiveProduct.color, 'Color')}
                                    className="w-5.5 h-5.5 bg-white border border-gray-250 hover:bg-gray-100 rounded-md flex items-center justify-center text-[10px] cursor-pointer shadow-sm"
                                  >
                                    📋
                                  </button>
                                  <button
                                    onClick={() => triggerMockAutofillField(simActiveProduct.color, 'color')}
                                    className="w-5.5 h-5.5 bg-gray-150 hover:bg-[#E85D24] text-gray-600 hover:text-white rounded-md flex items-center justify-center text-[10px] cursor-pointer transition-colors"
                                  >
                                    ⚡
                                  </button>
                                </div>
                              </div>
                            </div>
                          </div>

                          {/* Base Unit Price */}
                          <div className="border border-[#E85D24]/20 rounded-xl p-3 bg-amber-50/40 flex justify-between items-center shadow-sm">
                            <div>
                              <label className="text-[8.5px] uppercase tracking-wider font-extrabold text-[#E85D24] block">Base Price</label>
                              <span className="text-xs font-black text-gray-800 font-mono">PKR {simActiveProduct.basePrice}</span>
                            </div>
                            <div className="flex gap-1.5">
                              <button
                                onClick={() => triggerMockCopy(`PKR ${simActiveProduct.basePrice}`, 'Base Price')}
                                className="px-2.5 py-1.5 bg-white hover:bg-gray-100 border border-gray-200 rounded-lg text-[10px] font-bold text-gray-700 cursor-pointer transition-colors shadow-sm"
                              >
                                📋 Copy
                              </button>
                              <button
                                onClick={() => triggerMockAutofillField(simActiveProduct.basePrice, 'basePrice')}
                                className="px-2.5 py-1.5 bg-gray-800 hover:bg-black text-white rounded-lg text-[10px] font-bold cursor-pointer transition-colors shadow-sm"
                              >
                                ⚡ Fill
                              </button>
                            </div>
                          </div>

                          {/* Segment header: Variations prices */}
                          <div className="bg-gray-50 border border-gray-150 px-3 py-1.5 rounded-lg text-[9px] font-bold text-gray-500 uppercase tracking-wider font-display">
                            PACK VARIATIONS PACKS (EXCEL ROWS merged)
                          </div>

                          <div className="flex flex-col gap-1.5">
                            {simActiveProduct.variations.map((v, i) => (
                              <div key={v.pack} className="border border-gray-150 rounded p-2 bg-white">
                                <div className="text-[10.5px] font-bold text-gray-900 border-b border-gray-50 pb-1 mb-1">
                                  {v.pack}
                                </div>
                                <div className="grid grid-cols-2 gap-4">
                                  <div>
                                    <span className="text-[8px] text-gray-400 block">Actual Price</span>
                                    <div className="flex gap-1.5 items-center">
                                      <span className="text-xs font-mono font-bold">PKR {v.actualPrice}</span>
                                      <button
                                        onClick={() => triggerMockCopy(`PKR ${v.actualPrice}`, `${v.pack} Price`)}
                                        className="p-0.5 border border-gray-200 rounded text-[9px]"
                                        title="Copy price text"
                                      >
                                        📋
                                      </button>
                                    </div>
                                  </div>

                                  <div>
                                    <span className="text-[8px] text-[#E85D24] block">Special Sale price</span>
                                    <div className="flex gap-1.5 items-center justify-between">
                                      <span className="text-xs font-mono font-bold text-[#E85D24]">PKR {v.discountedPrice}</span>
                                      <div className="flex gap-1">
                                        <button
                                          onClick={() => triggerMockCopy(`PKR ${v.discountedPrice}`, `${v.pack} Sale`)}
                                          className="p-0.5 border border-[#E85D24]/10 rounded text-[9px]"
                                        >
                                          📋
                                        </button>
                                      </div>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            ))}
                          </div>

                          {/* Segment header: Logistics calculated metrics */}
                          <div className="bg-gray-100 px-2 py-1 rounded text-[9px] font-bold text-gray-600 tracking-wider flex justify-between items-center">
                            <span>AUTO WAREHOUSE LOGISTICS (FEATURE 3)</span>
                            <span className="text-[8px] text-gray-400 font-normal">Calculated based on first Pack pricing</span>
                          </div>

                          <div className="border border-gray-200 rounded p-2 bg-gray-50 flex flex-col gap-2">
                            
                            {/* Standard output measurements */}
                            <div className="flex justify-between items-start">
                              <div>
                                <label className="text-[8px] text-gray-400 font-bold block">Package L x W x H</label>
                                <span className="text-xs font-bold text-gray-800">
                                  {getSimulatedDimensions(simActiveProduct).length} x {getSimulatedDimensions(simActiveProduct).width} x {getSimulatedDimensions(simActiveProduct).height} cm
                                </span>
                              </div>
                              <div className="flex gap-1">
                                <button
                                  onClick={() => {
                                    const d = getSimulatedDimensions(simActiveProduct);
                                    triggerMockCopy(`${d.length}x${d.width}x${d.height}`, 'Dimensions');
                                  }}
                                  className="p-1 border border-gray-200 rounded text-[10px]"
                                >
                                  📋
                                </button>
                              </div>
                            </div>

                            {/* Standard output weight */}
                            <div className="flex justify-between items-start">
                              <div>
                                <label className="text-[8px] text-gray-400 font-bold block">Package Weight (grams)</label>
                                <span className="text-xs font-bold text-gray-800">
                                  {getSimulatedDimensions(simActiveProduct).weight}
                                </span>
                              </div>
                              <div className="flex gap-1">
                                <button
                                  onClick={() => triggerMockCopy(getSimulatedDimensions(simActiveProduct).weight, 'Weight')}
                                  className="p-1 border border-gray-200 rounded text-[10px]"
                                >
                                  📋
                                </button>
                                <button
                                  onClick={() => triggerMockAutofillField(getSimulatedDimensions(simActiveProduct).weight, 'weight')}
                                  className="p-1 bg-gray-100 hover:bg-[#E85D24] text-gray-600 hover:text-white rounded text-[10px]"
                                >
                                  ⚡
                                </button>
                              </div>
                            </div>

                            {/* Split input builders for filling form individuals */}
                            <div className="grid grid-cols-3 gap-1 border-t border-gray-150 pt-2 text-[10px]">
                              <div className="bg-white p-1 rounded border border-gray-200 text-center">
                                <label className="text-[7.5px] text-gray-400 block font-bold">L (cm)</label>
                                <div className="font-extrabold mt-0.5">{getSimulatedDimensions(simActiveProduct).length}</div>
                                <button 
                                  onClick={() => triggerMockAutofillField(getSimulatedDimensions(simActiveProduct).length, 'dimL')}
                                  className="mt-1 w-full bg-gray-100 text-[8px] py-0.5 rounded text-gray-500 hover:bg-[#E85D24] hover:text-white transition-colors"
                                >
                                  ⚡ Inject
                                </button>
                              </div>
                              <div className="bg-white p-1 rounded border border-gray-200 text-center">
                                <label className="text-[7.5px] text-gray-400 block font-bold">W (cm)</label>
                                <div className="font-extrabold mt-0.5">{getSimulatedDimensions(simActiveProduct).width}</div>
                                <button 
                                  onClick={() => triggerMockAutofillField(getSimulatedDimensions(simActiveProduct).width, 'dimW')}
                                  className="mt-1 w-full bg-gray-100 text-[8px] py-0.5 rounded text-gray-500 hover:bg-[#E85D24] hover:text-white transition-colors"
                                >
                                  ⚡ Inject
                                </button>
                              </div>
                              <div className="bg-white p-1 rounded border border-gray-200 text-center">
                                <label className="text-[7.5px] text-gray-400 block font-bold">H (cm)</label>
                                <div className="font-extrabold mt-0.5">{getSimulatedDimensions(simActiveProduct).height}</div>
                                <button 
                                  onClick={() => triggerMockAutofillField(getSimulatedDimensions(simActiveProduct).height, 'dimH')}
                                  className="mt-1 w-full bg-gray-100 text-[8px] py-0.5 rounded text-gray-500 hover:bg-[#E85D24] hover:text-white transition-colors"
                                >
                                  ⚡ Inject
                                </button>
                              </div>
                            </div>

                          </div>

                          {/* Warranty element */}
                          <div className="border border-gray-100 rounded p-1.5 bg-gray-50">
                            <label className="text-[8.5px] uppercase tracking-wider font-bold text-gray-400 block">Warranty Type</label>
                            <div className="flex justify-between items-center mt-0.5">
                              <span className="text-xs text-gray-800 font-extrabold">Seller Warranty — 1 Month</span>
                              <div className="flex gap-1">
                                <button
                                  onClick={() => triggerMockCopy('Seller Warranty — 1 Month', 'Warranty')}
                                  className="p-1 border border-gray-150 rounded text-[10px]"
                                >
                                  📋
                                </button>
                                <button
                                  onClick={() => triggerMockAutofillField('Seller Warranty — 1 Month', 'warranty')}
                                  className="p-1 bg-gray-100 hover:bg-[#E85D24] text-gray-600 hover:text-white rounded text-[10px]"
                                >
                                  ⚡
                                </button>
                              </div>
                            </div>
                          </div>

                          {/* Return Policy element */}
                          <div className="border border-gray-100 rounded p-1.5 bg-gray-50">
                            <div className="flex justify-between items-center mb-1">
                              <label className="text-[8.5px] uppercase tracking-wider font-bold text-gray-400 block">Return Policy Text</label>
                              <div className="flex gap-1">
                                <button
                                  onClick={() => {
                                    const text = `LIFEINNO Return Policy
• Product must be unused and in original condition
• Damaged or burnt items are not accepted
• Original packaging required
• Parcel opening video required
Contact us: 03357714860 or IM Chat.`;
                                    triggerMockCopy(text, 'Return Policy');
                                  }}
                                  className="p-1 border border-gray-150 rounded text-[10px] bg-white flex items-center gap-1 font-bold"
                                >
                                  📋 Copy Text
                                </button>
                                <button
                                  onClick={() => {
                                    const text = `LIFEINNO Return Policy
• Product must be unused and in original condition
• Damaged or burnt items are not accepted
• Original packaging required
• Parcel opening video required
Contact us: 03357714860 or IM Chat.`;
                                    triggerMockAutofillField(text, 'returnPolicy');
                                  }}
                                  className="p-1 bg-gray-100 hover:bg-[#E85D24] text-gray-600 hover:text-white rounded text-[10px]"
                                >
                                  ⚡
                                </button>
                              </div>
                            </div>
                            <span className="text-[9.5px] text-gray-500 font-mono block leading-relaxed max-height-[80px] overflow-y-auto border-l-2 border-[#E85D24] pl-2 py-0.5">
                              LIFEINNO Return Policy<br/>
                              • Product must be unused and in original condition<br/>
                              • Damaged or burnt items are not accepted<br/>
                              • Original packaging required<br/>
                              • Parcel opening video required<br/>
                              Contact us: 03357714860 or IM Chat.
                            </span>
                          </div>

                        </div>
                      </div>
                    )}

                  </div>

                  {/* VIRTUAL ACTIVE EXTENSION TOAST NOTIFIER */}
                  {simToast && (
                    <div className="absolute bottom-4 left-1/2 transform -translate-x-1/2 bg-gray-900 border border-[#E85D24] text-white px-3 py-1.5 rounded-full text-[10px] font-bold shadow-lg flex items-center gap-1 z-50 animate-bounce">
                      <Sparkles className="w-3 h-3 text-[#E85D24]" />
                      <span>{simToast}</span>
                    </div>
                  )}

                  {/* Virtual Chrome Frame Bottom Rim */}
                  <div className="bg-gray-100 text-gray-400 py-1 text-center text-[9px] font-mono border-t border-gray-205 select-none">
                    Status: Offline Local Caching OK
                  </div>
                </div>

              </div>

              {/* MOCK WEBSITE INTERACTIVE BROWSER (seller.daraz.pk) */}
              <div className="lg:col-span-8 flex flex-col">

                {/* Simulated URL bar and browser frame controls */}
                <div className="bg-slate-100 border border-slate-200 rounded-t-2xl px-5 py-3 flex items-center gap-3">
                  <div className="flex gap-1.5 shrink-0">
                    <span className="w-3 h-3 bg-red-400 rounded-full inline-block"></span>
                    <span className="w-3 h-3 bg-yellow-400 rounded-full inline-block"></span>
                    <span className="w-3 h-3 bg-green-400 rounded-full inline-block"></span>
                  </div>

                  <div className="flex-1 bg-white text-xs px-3 py-1.5 rounded-lg border border-slate-200 text-slate-500 font-mono flex items-center justify-between shadow-inner">
                    <span className="truncate text-[11px]">https://seller.daraz.pk/portal/product/publish?lang=en_PK</span>
                    <span className="text-[9px] text-emerald-600 font-extrabold bg-emerald-50 px-1.5 py-0.5 rounded border border-emerald-200 shrink-0">
                      SECURE SSL
                    </span>
                  </div>
                </div>

                {/* SIMULATED WEB PAGE CANVAS */}
                <div className="bg-white border-x border-b border-slate-200 rounded-b-2xl p-5 md:p-6 text-slate-700 min-h-[500px] shadow-sm">
                  
                  {/* Web page nested top brand bar */}
                  <div className="flex justify-between items-center border-b border-slate-100 pb-3 mb-6">
                    <div className="flex items-center gap-2">
                      <div className="w-7 h-7 bg-[#E85D24] rounded-lg text-white font-black text-center text-sm flex items-center justify-center shadow-sm">D</div>
                      <span className="text-xs font-black text-slate-900 font-display tracking-wider">DARAZ Seller Center <span className="text-gray-400 font-normal">| Pakistan</span></span>
                    </div>
                    <div className="text-[11px] text-slate-500 flex items-center gap-2 font-mono">
                      <span>Store ID: <strong className="text-slate-900">LIFEINNO_PK_M4</strong></span>
                      <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
                    </div>
                  </div>

                  {/* Browser simulated container splits */}
                  <div className="grid grid-cols-1 md:grid-cols-12 gap-5 pointer-events-auto">
                    
                    {/* Fake side Navigation rail */}
                    <div className="md:col-span-3 border-r border-slate-100 pr-4 flex flex-col gap-1.5 text-xs text-slate-500 select-none hidden md:flex">
                      <div className="text-[9px] font-extrabold text-slate-400 uppercase tracking-widest px-2 mb-1">Store Operations</div>
                      <div className="bg-[#FFF5F2] text-[#E85D24] font-extrabold px-3 py-2 rounded-lg border-l-2 border-[#E85D24]">📄 Add New Product</div>
                      <div className="hover:bg-slate-50 hover:text-slate-900 px-3 py-2 rounded-lg cursor-not-allowed transition-colors">📦 Product Catalog</div>
                      <div className="hover:bg-slate-50 hover:text-slate-900 px-3 py-2 rounded-lg cursor-not-allowed transition-colors">🛒 Orders & Reviews</div>
                      <div className="hover:bg-slate-50 hover:text-slate-900 px-3 py-2 rounded-lg cursor-not-allowed transition-colors">📊 Listing Analytics</div>
                      <div className="hover:bg-slate-50 hover:text-slate-900 px-3 py-2 rounded-lg cursor-not-allowed transition-colors">💬 店 IM Seller Chat</div>
                    </div>

                    {/* MOCK ADD PRODUCT SPECIFIC FORM */}
                    <div className="md:col-span-9 flex flex-col gap-5">
                      
                      <div className="flex justify-between items-center bg-slate-50 p-4 rounded-xl border border-slate-150">
                        <div>
                          <h4 className="text-sm font-black text-slate-800 font-display">Create a New Product Listing</h4>
                          <p className="text-xs text-slate-500 mt-1">Please provide warehouse descriptors and prices in Pakistani Rupee (PKR).</p>
                        </div>
                        <span className="text-[10px] text-[#E85D24] bg-[#E85D24]/10 border border-[#E85D24]/20 px-2 py-0.5 rounded font-bold font-mono shrink-0">STORYBOARD PREVIEW</span>
                      </div>

                      {/* Actual fields mapped onto Daraz standard schema inputs */}
                      <form className="flex flex-col gap-4 text-xs" onSubmit={(e) => e.preventDefault()}>
                                 {/* INPUT 1: TITLE */}
                        <div className="relative">
                          <label className="text-slate-600 block mb-1.5 font-bold text-[11px]">Product Name *</label>
                          <input
                            type="text"
                            placeholder="Enter professional listing title..."
                            value={darazFormState.title}
                            onFocus={() => setFocusedField('title')}
                            onChange={(e) => setDarazFormState(prev => ({ ...prev, title: e.target.value }))}
                            className={`w-full bg-slate-50 border px-3 py-2 rounded-xl text-slate-800 font-sans text-xs focus:bg-white focus:ring-1 focus:ring-[#E85D24] outline-none transition-all ${
                              focusedField === 'title' ? 'border-[#E85D24] duration-150' : 'border-slate-200'
                            } ${sparkleField === 'title' ? 'bg-[#FFF5F2] border-[#E85D24] shadow' : ''}`}
                          />
                          
                          {/* FLOATING ACTION FILL FROM CACHE TRIGGER (FEATURE 5 SIMULATION!) */}
                          {focusedField === 'title' && (
                            <div className="absolute right-3 top-7 z-10 animate-fade-in">
                              <button
                                type="button"
                                onClick={() => triggerMockAutofillField(simLastCopiedValue, 'title')}
                                className="bg-[#E85D24] text-white px-2.5 py-1 rounded-lg text-[10px] font-bold hover:scale-105 active:scale-95 transition-all cursor-pointer shadow-md"
                              >
                                ⚡ Fill from LIFEINNO Helper
                              </button>
                            </div>
                          )}
                        </div>

                        {/* SPLIT INPUTS: SIZE AND COLOR */}
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          
                          {/* SIZE */}
                          <div className="relative">
                            <label className="text-slate-600 block mb-1.5 font-bold text-[11px]">Size / Dimensions *</label>
                            <input
                              type="text"
                              placeholder="e.g. 6x6 Inches"
                              value={darazFormState.size}
                              onFocus={() => setFocusedField('size')}
                              onChange={(e) => setDarazFormState(prev => ({ ...prev, size: e.target.value }))}
                              className={`w-full bg-slate-50 border px-3 py-2 rounded-xl text-slate-800 font-sans text-xs focus:bg-white focus:ring-1 focus:ring-[#E85D24] outline-none transition-all ${
                                focusedField === 'size' ? 'border-[#E85D24]' : 'border-slate-200'
                              } ${sparkleField === 'size' ? 'bg-[#FFF5F2] border-[#E85D24] shadow' : ''}`}
                            />
                            {focusedField === 'size' && (
                              <div className="absolute right-3 top-7.5 z-10">
                                <button
                                  type="button"
                                  onClick={() => triggerMockAutofillField(simLastCopiedValue, 'size')}
                                  className="bg-[#E85D24] text-white px-2 py-0.5 rounded text-[10px] font-bold cursor-pointer"
                                >
                                  ⚡ Fill
                                </button>
                              </div>
                            )}
                          </div>

                          {/* COLOR */}
                          <div className="relative">
                            <label className="text-slate-600 block mb-1.5 font-bold text-[11px]">Color Option *</label>
                            <input
                              type="text"
                              placeholder="e.g. Silver Trim"
                              value={darazFormState.color}
                              onFocus={() => setFocusedField('color')}
                              onChange={(e) => setDarazFormState(prev => ({ ...prev, color: e.target.value }))}
                              className={`w-full bg-slate-50 border px-3 py-2 rounded-xl text-slate-800 font-sans text-xs focus:bg-white focus:ring-1 focus:ring-[#E85D24] outline-none transition-all ${
                                focusedField === 'color' ? 'border-[#E85D24]' : 'border-slate-200'
                              } ${sparkleField === 'color' ? 'bg-[#FFF5F2] border-[#E85D24] shadow' : ''}`}
                            />
                            {focusedField === 'color' && (
                              <div className="absolute right-3 top-7.5 z-10">
                                <button
                                  type="button"
                                  onClick={() => triggerMockAutofillField(simLastCopiedValue, 'color')}
                                  className="bg-[#E85D24] text-white px-2 py-0.5 rounded text-[10px] font-bold cursor-pointer"
                                >
                                  ⚡ Fill
                                </button>
                              </div>
                            )}
                          </div>

                        </div>

                        {/* BASE MANUFACTURER PRICE AND WAREHOUSE WEIGHT */}
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          
                          {/* BASE PRICE */}
                          <div className="relative">
                            <label className="text-slate-600 block mb-1.5 font-bold text-[11px]">Base Unit Price (PKR) *</label>
                            <input
                              type="text"
                              placeholder="e.g. 50"
                              value={darazFormState.basePrice}
                              onFocus={() => setFocusedField('basePrice')}
                              onChange={(e) => setDarazFormState(prev => ({ ...prev, basePrice: e.target.value }))}
                              className={`w-full bg-slate-50 border px-3 py-2 rounded-xl text-slate-800 font-sans text-xs focus:bg-white focus:ring-1 focus:ring-[#E85D24] outline-none transition-all ${
                                focusedField === 'basePrice' ? 'border-[#E85D24]' : 'border-slate-200'
                              } ${sparkleField === 'basePrice' ? 'bg-[#FFF5F2] border-[#E85D24] shadow' : ''}`}
                            />
                            {focusedField === 'basePrice' && (
                              <div className="absolute right-3 top-7.5 z-10">
                                <button
                                  type="button"
                                  onClick={() => triggerMockAutofillField(simLastCopiedValue, 'basePrice')}
                                  className="bg-[#E85D24] text-white px-2 py-0.5 rounded text-[10px] font-bold cursor-pointer"
                                >
                                  ⚡ Fill
                                </button>
                              </div>
                            )}
                          </div>

                          {/* PACK WEIGHT */}
                          <div className="relative">
                            <label className="text-slate-600 block mb-1.5 font-bold text-[11px]">Package Weight (g/grams) *</label>
                            <input
                              type="text"
                              placeholder="e.g. 300 grams"
                              value={darazFormState.weight}
                              onFocus={() => setFocusedField('weight')}
                              onChange={(e) => setDarazFormState(prev => ({ ...prev, weight: e.target.value }))}
                              className={`w-full bg-slate-50 border px-3 py-2 rounded-xl text-slate-800 font-sans text-xs focus:bg-white focus:ring-1 focus:ring-[#E85D24] outline-none transition-all ${
                                  focusedField === 'weight' ? 'border-[#E85D24]' : 'border-slate-200'
                              } ${sparkleField === 'weight' ? 'bg-[#FFF5F2] border-[#E85D24] shadow' : ''}`}
                            />
                            {focusedField === 'weight' && (
                              <div className="absolute right-3 top-7.5 z-10">
                                <button
                                  type="button"
                                  onClick={() => triggerMockAutofillField(simLastCopiedValue, 'weight')}
                                  className="bg-[#E85D24] text-white px-2 py-0.5 rounded text-[10px] font-bold cursor-pointer"
                                >
                                  ⚡ Fill
                                </button>
                              </div>
                            )}
                          </div>

                        </div>

                        {/* SPLIT INDIVIDUAL AUTO LOGISTICS DIMENSIONS */}
                        <div>
                          <label className="text-slate-600 block mb-1.5 font-bold text-[11px]">Package Dimensions (cm) *</label>
                          <div className="grid grid-cols-3 gap-3">
                            <div className="relative">
                              <span className="text-[9.5px] text-slate-500 block">Length (cm)</span>
                              <input
                                type="text"
                                placeholder="L"
                                value={darazFormState.dimL}
                                onFocus={() => setFocusedField('dimL')}
                                onChange={(e) => setDarazFormState(prev => ({ ...prev, dimL: e.target.value }))}
                                className={`w-full bg-slate-50 border p-2 rounded text-slate-800 font-mono mt-0.5 focus:ring-1 focus:ring-[#E85D24] outline-none ${
                                  focusedField === 'dimL' ? 'border-[#E85D24]' : 'border-slate-200'
                                } ${sparkleField === 'dimL' ? 'bg-[#FFF5F2]' : ''}`}
                              />
                            </div>
                            <div className="relative">
                              <span className="text-[9.5px] text-slate-500 block">Width (cm)</span>
                              <input
                                type="text"
                                placeholder="W"
                                value={darazFormState.dimW}
                                onFocus={() => setFocusedField('dimW')}
                                onChange={(e) => setDarazFormState(prev => ({ ...prev, dimW: e.target.value }))}
                                className={`w-full bg-slate-50 border p-2 rounded text-slate-800 font-mono mt-0.5 focus:ring-1 focus:ring-[#E85D24] outline-none ${
                                  focusedField === 'dimW' ? 'border-[#E85D24]' : 'border-slate-200'
                                } ${sparkleField === 'dimW' ? 'bg-[#FFF5F2]' : ''}`}
                              />
                            </div>
                            <div className="relative">
                              <span className="text-[9.5px] text-slate-500 block">Height (cm)</span>
                              <input
                                type="text"
                                placeholder="H"
                                value={darazFormState.dimH}
                                onFocus={() => setFocusedField('dimH')}
                                onChange={(e) => setDarazFormState(prev => ({ ...prev, dimH: e.target.value }))}
                                className={`w-full bg-slate-50 border p-2 rounded text-slate-800 font-mono mt-0.5 focus:ring-1 focus:ring-[#E85D24] outline-none ${
                                  focusedField === 'dimH' ? 'border-[#E85D24]' : 'border-slate-200'
                                } ${sparkleField === 'dimH' ? 'bg-[#FFF5F2]' : ''}`}
                              />
                            </div>
                          </div>
                        </div>

                        {/* WARRANTY */}
                        <div className="relative">
                          <label className="text-slate-600 block mb-1.5 font-bold text-[11px]">Warranty Type *</label>
                          <input
                            type="text"
                            placeholder="e.g. No Warranty"
                            value={darazFormState.warranty}
                            onFocus={() => setFocusedField('warranty')}
                            onChange={(e) => setDarazFormState(prev => ({ ...prev, warranty: e.target.value }))}
                            className={`w-full bg-slate-50 border px-3 py-2 rounded-xl text-slate-800 font-sans text-xs focus:bg-white focus:ring-1 focus:ring-[#E85D24] outline-none transition-all ${
                              focusedField === 'warranty' ? 'border-[#E85D24]' : 'border-slate-200'
                            } ${sparkleField === 'warranty' ? 'bg-[#FFF5F2] border-[#E85D24] shadow' : ''}`}
                          />
                          {focusedField === 'warranty' && (
                            <div className="absolute right-3 top-7.5 z-10 font-mono">
                              <button
                                type="button"
                                onClick={() => triggerMockAutofillField(simLastCopiedValue, 'warranty')}
                                className="bg-[#E85D24] text-white px-2 py-0.5 rounded text-[10px] font-bold cursor-pointer"
                              >
                                ⚡ Fill
                              </button>
                            </div>
                          )}
                        </div>

                        {/* RETURN POLICY AREA */}
                        <div className="relative">
                          <label className="text-slate-600 block mb-1.5 font-bold text-[11px]">Return Policy Terms *</label>
                          <textarea
                            placeholder="State exact replacement limitations..."
                            value={darazFormState.returnPolicy}
                            onFocus={() => setFocusedField('returnPolicy')}
                            rows={4}
                            onChange={(e) => setDarazFormState(prev => ({ ...prev, returnPolicy: e.target.value }))}
                            className={`w-full bg-slate-50 border px-3 py-2 rounded-xl text-slate-800 font-sans text-xs focus:bg-white focus:ring-1 focus:ring-[#E85D24] outline-none transition-all ${
                              focusedField === 'returnPolicy' ? 'border-[#E85D24]' : 'border-slate-200'
                            } ${sparkleField === 'returnPolicy' ? 'bg-[#FFF5F2] border-[#E85D24] shadow' : ''}`}
                          />
                          {focusedField === 'returnPolicy' && (
                            <div className="absolute right-3 top-7.5 z-10 font-mono">
                              <button
                                type="button"
                                onClick={() => triggerMockAutofillField(simLastCopiedValue, 'returnPolicy')}
                                className="bg-[#E85D24] text-white px-2 py-0.5 rounded text-[10px] font-bold cursor-pointer"
                              >
                                ⚡ Fill
                              </button>
                            </div>
                          )}
                        </div>

                      </form>

                      {/* Mock Submit Action (Faux completion test) */}
                      <div className="text-right border-t border-gray-850 pt-3 mt-2">
                        <button
                          type="button"
                          onClick={() => {
                            if (!darazFormState.title) {
                              alert('Please fill out the Product Title inside the form first!');
                              return;
                            }
                            alert(`🎉 SUCCESS! Listing preview parsed:\n\nTitle: ${darazFormState.title}\nSize: ${darazFormState.size}\nLogistics LWH: ${darazFormState.dimL}x${darazFormState.dimW}x${darazFormState.dimH} cm\nWeight: ${darazFormState.weight}\nStatus: Upload ready.`);
                            // Reset state
                            setDarazFormState({
                              title: '', size: '', color: '', basePrice: '', dimL: '', dimW: '', dimH: '', weight: '', warranty: '', returnPolicy: '', variationPrices: {}
                            });
                          }}
                          className="bg-emerald-500 hover:bg-emerald-600 text-white font-bold px-6 py-2 rounded-lg text-xs cursor-pointer shadow-lg transition-transform hover:scale-105 active:scale-95"
                        >
                          Submit Listing to Daraz (Mock Verification)
                        </button>
                      </div>

                    </div>

                  </div>

                </div>

              </div>

            </div>

          </div>
        )}

        {/* ======================= TAB 2: EXCEL TEMPLATE MANAGEMENT ======================= */}
        {activeTab === 'excel' && (
          <div className="flex flex-col gap-6 max-w-5xl mx-auto py-2">
            
            <div className="bg-gray-900 border border-gray-800 rounded-xl p-6 flex flex-col md:flex-row justify-between items-center gap-6">
              <div className="flex-1 flex gap-4">
                <div className="w-12 h-12 bg-white/5 border border-gray-850 text-white flex items-center justify-center rounded-xl font-mono text-2xl font-black">
                  🗄️
                </div>
                <div>
                  <h3 className="font-bold text-white text-base flex items-center gap-2">
                    Standard Excel Layout Specification
                  </h3>
                  <p className="text-xs text-gray-400 max-w-2xl mt-1 leading-relaxed">
                    The LIFEINNO Daraz Helper reads worksheets with dynamic row spanning. This design groups multi-SKU packs securely! Products with up to 12 Pack sizes have empty/null rows for general fields (Columns A-D, F) while variation data spans Columns E, G & H completely.
                  </p>
                </div>
              </div>
              <button
                id="btn-download-excel"
                onClick={downloadSampleXLSX}
                className="bg-emerald-500 hover:bg-emerald-600 text-white font-bold text-xs py-3 px-5 rounded-lg flex items-center gap-1.5 transition-all shadow-md shrink-0 cursor-pointer"
              >
                <Download className="w-4 h-4" />
                Download Sample Excel (.xlsx)
              </button>
            </div>

            {/* Ingestion block */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-stretch">
              
              {/* Detailed specs table */}
              <div className="bg-gray-900 border border-gray-800 rounded-xl p-5 flex flex-col">
                <h4 className="font-bold text-white text-xs uppercase tracking-wider mb-3 flex items-center gap-1">
                  <Grid className="w-4 h-4 text-[#E85D24]" />
                  Expected Columns Mapping
                </h4>

                <div className="flex-1 overflow-x-auto text-[11px]">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-gray-800 text-gray-500 font-bold">
                        <th className="py-2 px-1">Col</th>
                        <th className="py-2 px-2">Header Name</th>
                        <th className="py-2 px-2">Role & Rules</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-800 text-gray-300">
                      <tr>
                        <td className="py-2 px-1 text-[#E85D24] font-bold">A</td>
                        <td className="py-2 px-2 font-mono text-white">Product Serial Number</td>
                        <td className="py-2 px-2 text-gray-400">Unique identifier. Numeric/Text. Only on Row 1.</td>
                      </tr>
                      <tr>
                        <td className="py-2 px-1 text-[#E85D24] font-bold">B</td>
                        <td className="py-2 px-2 font-mono text-white">Professional Title</td>
                        <td className="py-2 px-2 text-gray-400">Main search text and listing name. First row only.</td>
                      </tr>
                      <tr>
                        <td className="py-2 px-1 text-[#E85D24] font-bold">C</td>
                        <td className="py-2 px-2 font-mono text-white">Size / Dimensions</td>
                        <td className="py-2 px-2 text-gray-400">Product dimensions e.g. "4x4 Inches". First row only.</td>
                      </tr>
                      <tr>
                        <td className="py-2 px-1 text-[#E85D24] font-bold">D</td>
                        <td className="py-2 px-2 font-mono text-white">Color Options</td>
                        <td className="py-2 px-2 text-gray-400">Available shades e.g. "Royal Gold". First row only.</td>
                      </tr>
                      <tr>
                        <td className="py-2 px-1 text-emerald-400 font-bold">E</td>
                        <td className="py-2 px-2 font-mono text-white">Variation Slot Name</td>
                        <td className="py-2 px-2 text-gray-400 font-bold text-emerald-400">Pack size name. Required on all spanned rows.</td>
                      </tr>
                      <tr>
                        <td className="py-2 px-1 text-[#E85D24] font-bold">F</td>
                        <td className="py-2 px-2 font-mono text-white">Base Unit Price (PKR)</td>
                        <td className="py-2 px-2 text-gray-400">Base manufacturing cost. First row only.</td>
                      </tr>
                      <tr>
                        <td className="py-2 px-1 text-emerald-400 font-bold">G</td>
                        <td className="py-2 px-2 font-mono text-white">Actual Price (Daraz Cut)</td>
                        <td className="py-2 px-2 text-gray-400">Full commission cost. Required on all rows.</td>
                      </tr>
                      <tr>
                        <td className="py-2 px-1 text-emerald-400 font-bold">H</td>
                        <td className="py-2 px-2 font-mono text-white">Discounted Price (Special Price)</td>
                        <td className="py-2 px-2 text-gray-400">Sale listing price. Required on all rows.</td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                <div className="bg-black/30 border border-gray-800 p-3 mt-4 rounded-lg">
                  <span className="text-[10px] uppercase font-bold text-gray-400 block tracking-wider">Example spanning:</span>
                  <p className="text-[10.5px] text-gray-400 font-mono mt-1 leading-normal">
                    Row 1: [4051, "3D Hex stickers", "6x6", "Gold", "Pack of 4", 20, 160, 120]<br/>
                    Row 2: [null, null,  null,  null, "Pack of 8", null, 320, 240]<br/>
                    Row 3: [null, null,  null,  null, "Pack of 12", null, 480, 360]
                  </p>
                </div>
              </div>

              {/* Spreadsheets dropzone */}
              <div 
                ref={dropzoneRef}
                onDragOver={(e) => { e.preventDefault(); e.currentTarget.style.borderColor = "#E85D24"; }}
                onDragLeave={(e) => { e.currentTarget.style.borderColor = "#374151"; }}
                onDrop={onFileDrop}
                className="bg-gray-900 border-2 border-dashed border-gray-800 rounded-xl p-8 text-center flex flex-col justify-center items-center transition-colors"
              >
                <div className="p-4 bg-emerald-500/10 text-emerald-400 rounded-full mb-3">
                  <Upload className="w-8 h-8" />
                </div>
                <h4 className="font-bold text-white text-base">Drag and Drop Your Excel Sheet</h4>
                <p className="text-xs text-gray-400 max-w-xs mt-1 leading-normal">
                  Drop your `.xlsx` or `.xls` spreadsheet file here to parse and inspect the custom rows structure instantly!
                </p>

                <div className="mt-4 flex gap-3">
                  <button
                    onClick={() => document.getElementById('parent-file-input')?.click()}
                    className="bg-gray-800 hover:bg-gray-700 text-white text-xs font-bold py-2 px-4 rounded border border-gray-750 transition-colors cursor-pointer"
                  >
                    Select File Locally
                  </button>
                  <button
                    onClick={() => {
                      const prebaked = getPrebakedProducts();
                      setProducts(prebaked);
                      setLoadedFilename('Procedural Template (33 Products)');
                      alert('Refreshed workspace template successfully with standard 33 LIFEINNO items.');
                    }}
                    className="bg-gray-850 hover:bg-gray-800 text-gray-400 hover:text-white text-xs py-2 px-4 rounded transition-colors cursor-pointer"
                  >
                    Reset Template
                  </button>
                </div>
                <input 
                  type="file" 
                  id="parent-file-input" 
                  className="hidden" 
                  accept=".xlsx, .xls"
                  onChange={onFileInputChange} 
                />

                {loadedFilename && (
                  <div className="mt-6 bg-black/40 border border-gray-800 px-4 py-2.5 rounded-lg text-left text-xs max-w-sm flex items-center gap-2">
                    <CheckCircle className="w-5 h-5 text-emerald-400 shrink-0" />
                    <div>
                      <span className="font-bold text-white block">Active Database Loaded</span>
                      <span className="text-[10px] text-gray-500 truncate block max-w-[220px]">{loadedFilename}</span>
                    </div>
                  </div>
                )}
              </div>

            </div>

            {/* GRIDS INTERACTION REVIEW */}
            <div className="bg-gray-900 border border-gray-800 rounded-xl p-5">
              <h4 className="font-bold text-white text-sm mb-3">
                Active Parsed Structure Grid Review ({products.length} Products Found)
              </h4>

              <div className="border border-gray-850 rounded-lg overflow-hidden max-h-[350px] overflow-y-auto">
                <table className="w-full text-left text-xs border-collapse font-sans">
                  <thead>
                    <tr className="bg-black/60 border-b border-gray-850 text-gray-400 font-bold">
                      <th className="p-3">SN</th>
                      <th className="p-3">Name / Listing Title</th>
                      <th className="p-3">Dimensions</th>
                      <th className="p-3">Color</th>
                      <th className="p-3">Base (PKR)</th>
                      <th className="p-3">Variations Matrix (Pack : Cut Price / Sale Price)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-850 text-gray-300">
                    {products.map((p, idx) => (
                      <tr key={p.serial} className="hover:bg-gray-850/30">
                        <td className="p-3 font-mono font-bold text-[#E85D24]">#{p.serial}</td>
                        <td className="p-3 font-bold text-white max-w-[220px]">{p.title}</td>
                        <td className="p-3">{p.size}</td>
                        <td className="p-3 text-gray-400">{p.color}</td>
                        <td className="p-3 font-mono">PKR {p.basePrice}</td>
                        <td className="p-3">
                          <div className="flex flex-wrap gap-1">
                            {p.variations.map((v) => (
                              <span 
                                key={v.pack} 
                                className="bg-black/40 rounded px-1.5 py-0.5 text-[10px] border border-gray-800 text-gray-300"
                                title={`Actual Cut: PKR ${v.actualPrice} | Sale Price: PKR ${v.discountedPrice}`}
                              >
                                {v.pack} (<strong className="text-emerald-400">PKR {v.discountedPrice}</strong>)
                              </span>
                            ))}
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

          </div>
        )}

        {/* ======================= TAB 3: PACK EXPORTER & ZIP ASSEMBLY ======================= */}
        {activeTab === 'code' && (
          <div className="flex flex-col gap-6 max-w-5xl mx-auto py-2">
            
            {/* Download panel */}
            <div className="bg-gray-900 border border-gray-800 rounded-xl p-6 flex flex-col md:flex-row justify-between items-center gap-6 shadow-xl">
              <div className="flex gap-4">
                <div className="w-14 h-14 bg-[#E85D24]/10 border border-[#E85D24]/20 rounded-2xl flex items-center justify-center text-3xl">
                  🎁
                </div>
                <div>
                  <h3 className="font-extrabold text-white text-base">Bundled Chrome Extension Builder</h3>
                  <p className="text-xs text-gray-400 max-w-2xl mt-1 leading-relaxed">
                    Download the compiled Chrome Extension bundle. This zip archive integrates standard Manifest V3, SidePanel automation, content triggers, offline local caching, and is optimized to load unpacked instantly inside Google Chrome.
                  </p>
                </div>
              </div>

              <button
                id="btn-download-zip"
                onClick={handleDownloadZipOfExtension}
                className="bg-[#E85D24] hover:bg-[#d14f1b] text-white font-black text-xs py-3 px-6 rounded-lg flex items-center gap-2 transition-all shadow-lg text-center shrink-0 cursor-pointer animate-pulse hover:animate-none"
              >
                <Download className="w-4 h-4" />
                📥 Download Chrome Extension .ZIP
              </button>
            </div>

            {/* Side-by-side: Code directory tabs and step-by-step loading guide */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
              
              {/* CODE TAB REVIEWER */}
              <div className="lg:col-span-8 bg-gray-900 border border-gray-800 rounded-xl flex flex-col overflow-hidden">
                
                {/* Visual tabs heading */}
                <div className="bg-black/40 border-b border-gray-850 px-3 py-2 flex justify-between items-center select-none overflow-x-auto">
                  <span className="text-xs font-mono font-bold text-gray-400 flex items-center gap-1">
                    <FileCode className="w-4 h-4 text-[#E85D24]" />
                    Extension Files Explorer
                  </span>
                  
                  {/* Tab switches */}
                  <div className="flex gap-1 overflow-x-auto shrink-0 max-w-md md:max-w-xl py-1">
                    {[
                      { key: 'manifest', label: 'manifest.json' },
                      { key: 'popup_html', label: 'popup.html' },
                      { key: 'popup_js', label: 'popup.js' },
                      { key: 'sidebar_html', label: 'sidebar.html' },
                      { key: 'sidebar_js', label: 'sidebar.js' },
                      { key: 'content', label: 'content.js' },
                      { key: 'background', label: 'background.js' },
                      { key: 'styles', label: 'styles.css' },
                      { key: 'readme', label: 'README' }
                    ].map(tab => (
                      <button
                        key={tab.key}
                        id={`code-tab-${tab.key}`}
                        onClick={() => setSelectedCodeTab(tab.key)}
                        className={`px-2 py-1 rounded text-[10.5px] font-mono cursor-pointer transition-colors whitespace-nowrap ${
                          selectedCodeTab === tab.key
                            ? 'bg-gray-800 text-[#E85D24] font-bold border border-gray-750'
                            : 'text-gray-500 hover:text-white'
                        }`}
                      >
                        {tab.label}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Sub-header detailing chosen filename */}
                <div className="px-4 py-2 border-b border-gray-850 flex items-center justify-between text-xs bg-gray-950/20">
                  <span className="font-mono text-gray-350">
                    File: <strong className="text-white font-mono">{getCodeTabFilename()}</strong>
                  </span>
                  <button
                    onClick={handleCopyCodeText}
                    className="bg-gray-850 hover:bg-gray-800 border border-gray-750 text-gray-300 hover:text-white px-3 py-1 rounded text-[10px] font-bold flex items-center gap-1 cursor-pointer transition-colors"
                  >
                    {copyCodeSuccess ? (
                      <>
                        <Check className="w-3.5 h-3.5 text-emerald-400" />
                        <span>Copied!</span>
                      </>
                    ) : (
                      <>
                        <Copy className="w-3.5 h-3.5" />
                        <span>Copy Code Block</span>
                      </>
                    )}
                  </button>
                </div>

                {/* Visual code layout content */}
                <div className="bg-gray-950 p-4 overflow-x-auto max-h-[420px] overflow-y-auto">
                  <pre className="text-[11px] font-mono leading-relaxed text-emerald-400 font-normal">
                    <code>{getCodeTabString()}</code>
                  </pre>
                </div>

              </div>

              {/* CHROME DEVELOPMENT UNFOLD ROADMAP */}
              <div className="lg:col-span-4 bg-gray-900 border border-gray-800 rounded-xl p-5 flex flex-col gap-4">
                <h4 className="font-bold text-white text-xs uppercase tracking-wider mb-2 flex items-center gap-1.5">
                  <Chrome className="w-4 h-4 text-[#E85D24]" />
                  How to Load unpacked in Chrome
                </h4>

                <div className="flex flex-col gap-4 text-xs font-mono">
                  
                  {/* Step 1 */}
                  <div className="flex gap-3">
                    <div className="w-6 h-6 bg-gray-800 rounded-full flex items-center justify-center text-xs font-bold font-mono text-[#E85D24] shrink-0 border border-gray-750">
                      1
                    </div>
                    <div>
                      <span className="text-white font-bold block">Download & Extract</span>
                      <p className="text-[11px] text-gray-400 mt-0.5">Click the orange bundle button. Extract the downloaded `.zip` file into a dedicated local folder.</p>
                    </div>
                  </div>

                  {/* Step 2 */}
                  <div className="flex gap-3">
                    <div className="w-6 h-6 bg-gray-800 rounded-full flex items-center justify-center text-xs font-bold font-mono text-[#E85D24] shrink-0 border border-gray-750">
                      2
                    </div>
                    <div>
                      <span className="text-white font-bold block">Open Extensions Manager</span>
                      <p className="text-[11px] text-gray-400 mt-0.5">In a new Google Chrome browser tab, navigate directly to exactly: <code className="text-emerald-400">chrome://extensions/</code></p>
                    </div>
                  </div>

                  {/* Step 3 */}
                  <div className="flex gap-3">
                    <div className="w-6 h-6 bg-gray-800 rounded-full flex items-center justify-center text-xs font-bold font-mono text-[#E85D24] shrink-0 border border-gray-750">
                      3
                    </div>
                    <div>
                      <span className="text-white font-bold block">Toggle Developer Mode</span>
                      <p className="text-[11px] text-gray-400 mt-0.5">In the top right corner of the Extensions panel, toggle the **Developer Mode** switch to **ON**.</p>
                    </div>
                  </div>

                  {/* Step 4 */}
                  <div className="flex gap-3">
                    <div className="w-6 h-6 bg-gray-800 rounded-full flex items-center justify-center text-xs font-bold font-mono text-[#E85D24] shrink-0 border border-gray-750">
                      4
                    </div>
                    <div>
                      <span className="text-white font-bold block">Load Unpacked Folder</span>
                      <p className="text-[11px] text-gray-400 mt-0.5">Click the top-left button **"Load Unpacked"**. Select your extracted folder containing all extension code files.</p>
                    </div>
                  </div>

                </div>

                <div className="border border-white/5 bg-gray-950/40 rounded-xl p-3 mt-4 text-[11px] text-gray-400 flex items-start gap-2 leading-relaxed">
                  <span className="text-[#E85D24] font-bold">⚠️ SUPPORT FOR DARAZ PROPORTIONALITY:</span>
                  <p>
                    Once active, the extension runs of both product details and sidebar options automatically when visiting seller.daraz.pk listings.
                  </p>
                </div>
              </div>

            </div>

          </div>
        )}

      </main>

      {/* FOOTER METADATA BAR */}
      <footer className="bg-gray-950 border-t border-gray-900 px-6 py-4 text-center text-gray-500 text-xs flex flex-col sm:flex-row justify-between items-center gap-3">
        <div className="flex items-center gap-2 select-none">
          <span className="font-mono">Workspace Version: 1.0</span>
          <span>|</span>
          <span className="font-mono text-gray-600">Built for Mudassir Bashir</span>
        </div>
        <p className="font-mono">
          © 2026 LIFEINNO Corp. All rights reserved.
        </p>
      </footer>

    </div>
  );
}
