import React, { useState, useRef, useMemo, ChangeEvent, useEffect } from 'react';
import * as XLSX from 'xlsx';
// @ts-ignore
import html2pdf from 'html2pdf.js';
import { Upload, Printer, Trash2, FileSpreadsheet, Search, CheckCircle2, Download, ChevronLeft, Edit3, X } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { SampleData, CustomerContact, LabelData } from './types';

export default function App() {
  const [samples, setSamples] = useState<SampleData[]>(() => {
    const saved = localStorage.getItem('delivery-master-samples');
    if (!saved) return [];
    const parsed = JSON.parse(saved) as SampleData[];
    // Migration: ensure every sample has a unique ID
    const migrated = parsed.map(s => ({
      ...s,
      id: s.id || crypto.randomUUID()
    }));
    return migrated;
  });
  const [contacts, setContacts] = useState<CustomerContact[]>(() => {
    const saved = localStorage.getItem('delivery-master-contacts');
    return saved ? JSON.parse(saved) : [];
  });
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());

  // Save data whenever updated
  const updateContacts = (newContacts: CustomerContact[]) => {
    setContacts(newContacts);
    localStorage.setItem('delivery-master-contacts', JSON.stringify(newContacts));
  };

  const updateSamples = (newSamples: SampleData[]) => {
    setSamples(newSamples);
    localStorage.setItem('delivery-master-samples', JSON.stringify(newSamples));
  };

  const [searchTerm, setSearchTerm] = useState('');
  const [isPrinting, setIsPrinting] = useState(false);
  const [isGeneratingPDF, setIsGeneratingPDF] = useState(false);

  // Manual entry states
  const [manualSample, setManualSample] = useState({
    samplePo: '',
    piNo: '',
    customer: '',
    sampleType: ''
  });
  const [manualContact, setManualContact] = useState({
    customerName: '',
    contactPerson: '',
    phoneNumber: '',
    address: ''
  });

  // Editing states
  const [editingContactIndex, setEditingContactIndex] = useState<number | null>(null);
  const [editContactBuffer, setEditContactBuffer] = useState<CustomerContact | null>(null);

  // Sample editing states
  const [editingSampleId, setEditingSampleId] = useState<string | null>(null);
  const [editSampleBuffer, setEditSampleBuffer] = useState<SampleData | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const contactInputRef = useRef<HTMLInputElement>(null);
  const printRef = useRef<HTMLDivElement>(null);

  const generatePDF = () => {
    const element = printRef.current;
    if (!element) return;

    setIsGeneratingPDF(true);
    const opt = {
      margin: 0,
      filename: `Delivery_Labels_${new Date().getTime()}.pdf`,
      image: { type: 'jpeg' as const, quality: 0.98 },
      html2canvas: { 
        scale: 2, 
        useCORS: true,
        letterRendering: true,
        logging: false
      },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' as const }
    };

    html2pdf().set(opt).from(element).save().then(() => {
      setIsGeneratingPDF(false);
    });
  };

  // Auto-trigger PDF on enter print mode
  useEffect(() => {
    if (isPrinting) {
      const timer = setTimeout(() => {
        generatePDF();
      }, 1000);
      return () => clearTimeout(timer);
    }
  }, [isPrinting]);

  // Download Templates
  const downloadTemplate = (type: 'samples' | 'contacts') => {
    const wb = XLSX.utils.book_new();
    let ws;
    let fileName = '';

    if (type === 'samples') {
      const data = [
        ['SI NO', 'Sample PO', 'Po No', 'Pi No', 'Customer', 'Sample Type'],
        [1, '182/0136', 'PO/182/2026', 'MTLA/0136/2025', 'AST KNITWEAR LTD.', 'PRICE BARCODE']
      ];
      ws = XLSX.utils.aoa_to_sheet(data);
      fileName = 'Sample_PO_Template.xlsx';
    } else {
      const data = [
        ['Customer Name', 'Contact Person', 'Phone Number', 'Address'],
        ['AST KNITWEAR LTD.', 'Hannan | Executive,', 'm: +880 1329721679', 'Gazipur, Dhaka']
      ];
      ws = XLSX.utils.aoa_to_sheet(data);
      fileName = 'Contact_List_Template.xlsx';
    }

    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, fileName);
  };

  // Parse Excel Files
  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet) as any[];

      if (jsonData.length > 0) {
        const firstRow = jsonData[0];
        const keys = Object.keys(firstRow).map(k => k.toLowerCase().trim());
        
        if (keys.includes('customer name') || keys.includes('contact person')) {
          const newContacts: CustomerContact[] = jsonData
            .filter(item => (item['Customer Name'] || item['customer name']))
            .map(item => ({
              customerName: String(item['Customer Name'] || item['customer name'] || '').trim().replace(/^"/, ''),
              contactPerson: String(item['Contact Person'] || item['contact person'] || '').trim(),
              phoneNumber: String(item['Phone Number'] || item['phone number'] || '').trim(),
              address: String(item['Address'] || item['address'] || '').trim(),
            }));
          updateContacts(newContacts);
        } else {
          const newSamples: SampleData[] = jsonData
            .filter(item => (item['Sample PO'] || item['sample po'] || item['Customer'] || item['customer']))
            .map((item, idx) => ({
              id: crypto.randomUUID(),
              siNo: item['SI NO'] || item['si no'] || idx + 1,
              samplePo: String(item['Sample PO'] || item['sample po'] || '').trim().replace(/^"/, ''),
              poNo: String(item['Po No'] || item['po no'] || '').trim(),
              piNo: String(item['Pi No'] || item['pi no'] || '').trim().replace(/^"/, ''),
              customer: String(item['Customer'] || item['customer'] || '').trim().replace(/^"/, ''),
              sampleType: String(item['Sample Type'] || item['sample type'] || '').trim(),
            }));
          updateSamples(newSamples);
        }
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Merge Data for Labels
  const filteredSamples = useMemo(() => {
    return samples.filter(s => 
      s.customer.toLowerCase().includes(searchTerm.toLowerCase()) ||
      s.samplePo.toLowerCase().includes(searchTerm.toLowerCase())
    );
  }, [samples, searchTerm]);

  const selectedLabels: LabelData[] = useMemo(() => {
    const selectedSamples = samples.filter(s => selectedIds.has(s.id));
    const uniqueCustomers = new Map<string, LabelData>();

    selectedSamples.forEach(s => {
      const customerKey = s.customer.trim().toLowerCase();
      if (!uniqueCustomers.has(customerKey)) {
        const contact = contacts.find(c => 
          c.customerName.trim().toLowerCase() === customerKey
        );
        uniqueCustomers.set(customerKey, {
          ...s,
          contactPerson: contact?.contactPerson || 'N/A',
          phoneNumber: contact?.phoneNumber || 'N/A',
          address: contact?.address || ''
        });
      } else {
        const existing = uniqueCustomers.get(customerKey)!;
        // Concatenate if not already present
        if (!existing.samplePo.split(', ').includes(s.samplePo)) {
          existing.samplePo = existing.samplePo ? `${existing.samplePo}, ${s.samplePo}` : s.samplePo;
        }
        if (!existing.piNo.split(', ').includes(s.piNo)) {
          existing.piNo = existing.piNo ? `${existing.piNo}, ${s.piNo}` : s.piNo;
        }
      }
    });

    return Array.from(uniqueCustomers.values());
  }, [samples, selectedIds, contacts]);

  const toggleSelect = (id: string) => {
    const newSelected = new Set(selectedIds);
    if (newSelected.has(id)) newSelected.delete(id);
    else newSelected.add(id);
    setSelectedIds(newSelected);
  };

  const selectAll = () => {
    if (selectedIds.size === filteredSamples.length) {
      setSelectedIds(new Set());
    } else {
      const allIds = filteredSamples.map(s => s.id);
      setSelectedIds(new Set(allIds));
    }
  };

  const clearData = () => {
    updateSamples([]);
    setSelectedIds(new Set());
  };

  // Handle Paste Data
  const handlePaste = (e: React.ClipboardEvent) => {
    const pasteData = e.clipboardData.getData('text');
    if (!pasteData) return;

    // Robust TSV Parser that respects quotes and newlines inside cells
    const parseTsv = (text: string) => {
      const rows: string[][] = [];
      let currentRow: string[] = [];
      let currentField = '';
      let insideQuotes = false;
      
      for (let i = 0; i < text.length; i++) {
        const char = text[i];
        const nextChar = text[i + 1];

        if (char === '"' && insideQuotes && nextChar === '"') {
          currentField += '"';
          i++;
        } else if (char === '"') {
          insideQuotes = !insideQuotes;
          // Don't add the quote character itself to the field
        } else if (char === '\t' && !insideQuotes) {
          currentRow.push(currentField);
          currentField = '';
        } else if ((char === '\r' || char === '\n') && !insideQuotes) {
          if (char === '\r' && nextChar === '\n') i++;
          currentRow.push(currentField);
          rows.push(currentRow);
          currentRow = [];
          currentField = '';
        } else {
          currentField += char;
        }
      }
      if (currentField || currentRow.length > 0) {
        currentRow.push(currentField);
        rows.push(currentRow);
      }
      return rows;
    };

    const allRows = parseTsv(pasteData).filter(row => row.some(cell => cell.trim() !== ''));
    if (allRows.length === 0) return;

    const headers = allRows[0].map(h => h.toLowerCase().trim());
    const dataRows = allRows.slice(1);

    const hasContactHeaders = headers.includes('customer name') || headers.includes('contact person');
    
    if (hasContactHeaders) {
      // It's a Contact List
      const newContacts: CustomerContact[] = dataRows
        .map(cells => {
          const getVal = (search: string) => {
            const idx = headers.indexOf(search);
            return idx !== -1 ? (cells[idx] || '').trim() : '';
          };
          return {
            customerName: getVal('customer name') || (cells[0] || '').trim(),
            contactPerson: getVal('contact person') || (cells[1] || '').trim(),
            phoneNumber: getVal('phone number') || (cells[2] || '').trim(),
            address: getVal('address') || (cells[3] || '').trim(),
          };
        })
        .filter(c => c.customerName !== '');

      updateContacts([...contacts, ...newContacts]);
    } else {
      // It's a Sample List
      const newSamples: SampleData[] = dataRows
        .map((cells, idx) => {
          const getVal = (search: string) => {
            const idx = headers.indexOf(search);
            return idx !== -1 ? (cells[idx] || '').trim() : '';
          };
          return {
            id: crypto.randomUUID(),
            siNo: getVal('si no') || idx + samples.length + 1,
            samplePo: getVal('sample po') || (cells[1] || '').trim(),
            poNo: getVal('po no') || (cells[2] || '').trim(),
            piNo: getVal('pi no') || (cells[3] || '').trim(),
            customer: getVal('customer') || (cells[4] || '').trim(),
            sampleType: getVal('sample type') || (cells[5] || '').trim(),
          };
        })
        .filter(s => s.samplePo !== '' || s.customer !== '');

      updateSamples([...samples, ...newSamples]);
    }
  };

  const [isDragging, setIsDragging] = useState(false);
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };
  const handleDragLeave = () => setIsDragging(false);
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      const fakeEvent = { target: { files: [file] } } as any;
      handleFileUpload(fakeEvent);
    }
  };

  const addManualSample = () => {
    if (!manualSample.samplePo || !manualSample.customer) {
      alert('Please fill at least Sample PO and Customer');
      return;
    }
    const newSample: SampleData = {
      id: crypto.randomUUID(),
      siNo: samples.length + 1,
      ...manualSample,
      poNo: manualSample.samplePo // defaulting for consistency
    };
    updateSamples([...samples, newSample]);
    setManualSample({ samplePo: '', piNo: '', customer: '', sampleType: '' });
  };

  const handleManualSamplePaste = (e: React.ClipboardEvent) => {
    const data = e.clipboardData.getData('text');
    if (data.includes('\t')) {
      e.preventDefault();
      const cells = data.split('\t').map(c => c.trim());
      // Try to intelligently map cells
      setManualSample({
        samplePo: cells[0] || '',
        piNo: cells[1] || '',
        customer: cells[2] || '',
        sampleType: cells[3] || ''
      });
    }
  };

  const handleManualContactPaste = (e: React.ClipboardEvent) => {
    const data = e.clipboardData.getData('text');
    if (data.includes('\t')) {
      e.preventDefault();
      const cells = data.split('\t').map(c => c.trim());
      setManualContact({
        customerName: cells[0] || '',
        contactPerson: cells[1] || '',
        phoneNumber: cells[2] || '',
        address: cells[3] || ''
      });
    }
  };

  const addManualContact = () => {
    if (!manualContact.customerName || !manualContact.contactPerson) {
      alert('Please fill Customer Name and Contact Person');
      return;
    }
    updateContacts([...contacts, manualContact]);
    setManualContact({ customerName: '', contactPerson: '', phoneNumber: '', address: '' });
  };

  const startEditingContact = (index: number) => {
    setEditingContactIndex(index);
    setEditContactBuffer({ ...contacts[index] });
  };

  const saveEditedContact = () => {
    if (!editContactBuffer || editingContactIndex === null) return;
    const newContacts = [...contacts];
    newContacts[editingContactIndex] = editContactBuffer;
    updateContacts(newContacts);
    setEditingContactIndex(null);
    setEditContactBuffer(null);
  };

  const cancelEditingContact = () => {
    setEditingContactIndex(null);
    setEditContactBuffer(null);
  };

  const deleteContact = (index: number) => {
    if (confirm('Are you sure you want to delete this contact?')) {
      const newContacts = contacts.filter((_, i) => i !== index);
      updateContacts(newContacts);
    }
  };

  const startEditingSample = (sample: SampleData) => {
    setEditingSampleId(sample.id);
    setEditSampleBuffer({ ...sample });
  };

  const saveEditedSample = () => {
    if (!editSampleBuffer || !editingSampleId) return;
    const newSamples = samples.map(s => s.id === editingSampleId ? editSampleBuffer : s);
    updateSamples(newSamples);
    setEditingSampleId(null);
    setEditSampleBuffer(null);
  };

  const cancelEditingSample = () => {
    setEditingSampleId(null);
    setEditSampleBuffer(null);
  };

  const deleteSample = (id: string) => {
    if (confirm('Are you sure you want to delete this sample?')) {
      const newSamples = samples.filter(s => s.id !== id);
      updateSamples(newSamples);
    }
  };

  if (isPrinting) {
    return (
      <div className="bg-neutral-800 min-h-screen text-black">
        <div className="p-4 no-print flex justify-between bg-white/10 backdrop-blur-md items-center sticky top-0 z-50 shadow-2xl">
          <div className="flex items-center gap-4">
            <button 
              onClick={() => setIsPrinting(false)}
              className="flex items-center gap-2 px-4 py-2 bg-white/10 border border-white/20 text-white rounded-lg hover:bg-white/20 transition-all font-medium text-sm"
            >
              <ChevronLeft size={18} />
              Back to Dashboard
            </button>
            <div className="h-6 w-px bg-white/20"></div>
            <span className="text-white font-bold text-sm">
              {selectedLabels.length} {selectedLabels.length === 1 ? 'Label' : 'Labels'} Organized
            </span>
          </div>

          <div className="flex items-center gap-3">
            <button 
              onClick={generatePDF}
              disabled={isGeneratingPDF}
              className="px-6 py-2 bg-blue-600 text-white rounded-lg flex items-center gap-2 hover:bg-blue-500 transition-all font-bold text-sm shadow-lg shadow-blue-600/20 disabled:opacity-50"
            >
              {isGeneratingPDF ? (
                <>
                  <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin"></div>
                  Generating PDF...
                </>
              ) : (
                <>
                  <Download size={18} />
                  Download PDF
                </>
              )}
            </button>
            <button 
              onClick={() => window.print()}
              disabled={isGeneratingPDF}
              className="px-6 py-2 bg-white text-neutral-900 rounded-lg flex items-center gap-2 hover:bg-neutral-100 transition-all font-bold text-sm shadow-xl"
            >
              <Printer size={18} />
              System Print
            </button>
          </div>
        </div>

        {/* Print Layout */}
        <div className="flex flex-col items-center py-12 px-4 overflow-x-auto min-w-[210mm]">
          <div 
            ref={printRef}
            className="print-container p-0 bg-white shadow-[0_0_100px_rgba(0,0,0,0.5)] w-[210mm] min-h-[297mm]"
          >
            <div className="grid grid-cols-2 gap-0 border-collapse">
            {selectedLabels.map((label, idx) => (
              <div 
                key={idx} 
                className="label-box h-[74.25mm] w-[105mm] border border-gray-300 p-4 flex flex-col justify-between overflow-hidden"
              >
                <div className="flex flex-col flex-1">
                  <div className="flex justify-between items-start mb-2 border-b-2 border-neutral-900 pb-1">
                    <div className="flex flex-col">
                      <div className="text-[7px] text-neutral-400 font-black uppercase tracking-widest">From:</div>
                      <div className="text-[9px] font-black text-neutral-900 leading-tight">MAINETTI TLA SOLUTIONS BANGLADESH PVT. LTD</div>
                    </div>
                    <div className="bg-neutral-900 text-white text-[8px] font-mono px-2 py-0.5 rounded tracking-tighter">BATCH #0{idx+1}</div>
                  </div>

                  <div className="flex justify-between items-center mb-1">
                    <h2 className="text-[13px] font-black uppercase tracking-tight text-neutral-800 border-b border-neutral-300 w-fit pb-0.5">Sample Delivery</h2>
                  </div>
                  
                  <div className="grid grid-cols-2 gap-x-4 gap-y-1 mt-1">
                    <div className="flex flex-col">
                       <span className="text-neutral-400 text-[7px] font-black uppercase tracking-tighter">Principal / PI No:</span>
                       <span className="text-[10px] font-bold leading-[1.2] text-neutral-800 break-words line-clamp-3 overflow-hidden h-[3.8em]">{label.piNo}</span>
                    </div>
                    <div className="flex flex-col text-right">
                       <span className="text-neutral-400 text-[7px] font-black uppercase tracking-tighter">Sample Type:</span>
                       <div className="mt-1">
                        <span className="inline-block bg-neutral-100 text-neutral-700 text-[8px] px-1.5 py-0.5 rounded font-black italic border border-neutral-200">{label.sampleType}</span>
                       </div>
                    </div>
                  </div>
                </div>

                <div className="mt-auto pt-4 border-t-2 border-neutral-900">
                  <div className="text-[7px] font-black text-neutral-400 uppercase tracking-[0.2em] mb-1">Deliver To</div>
                  <div className="space-y-1">
                    <div className="text-[20px] font-black text-blue-950 leading-[0.95] uppercase break-words line-clamp-2 min-h-[1.95em] flex items-center">{label.customer}</div>
                    {label.address && (
                      <div className="text-[9px] font-bold text-neutral-500 leading-tight border-l-2 border-neutral-300 pl-2 mt-1 line-clamp-1 italic">
                        {label.address}
                      </div>
                    )}
                    <div className="flex items-center justify-between pt-1">
                      <div className="text-[12px] font-black flex items-baseline gap-1 text-neutral-900">
                        <span className="text-neutral-400 text-[8px] font-bold uppercase tracking-tighter">Attn:</span> {label.contactPerson}
                      </div>
                      <div className="text-[10px] text-neutral-500 font-mono font-bold bg-neutral-100 px-2 py-0.5 rounded border border-neutral-200">{label.phoneNumber}</div>
                    </div>
                  </div>
                </div>
              </div>
            ))}
            </div>
          </div>
        </div>

        <style>{`
          @media print {
            body { margin: 0; padding: 0; }
            .no-print { display: none !important; }
            .print-container { width: 210mm; margin: 0; }
            .label-box { 
              page-break-inside: avoid;
              border: 0.1mm solid #ddd !important;
              background-color: white !important;
              color: #171717 !important;
            }
          }
          /* Fix for html2canvas / html2pdf not supporting oklch (Tailwind v4) */
          .print-container {
            font-family: sans-serif;
            background-color: #ffffff !important;
            --color-white: #ffffff !important;
            --color-black: #000000 !important;
            --color-neutral-50: #fafafa !important;
            --color-neutral-100: #f5f5f5 !important;
            --color-neutral-200: #e5e5e5 !important;
            --color-neutral-300: #d4d4d4 !important;
            --color-neutral-400: #a3a3a3 !important;
            --color-neutral-500: #737373 !important;
            --color-neutral-600: #525252 !important;
            --color-neutral-700: #404040 !important;
            --color-neutral-800: #262626 !important;
            --color-neutral-900: #171717 !important;
            --color-neutral-950: #0a0a0a !important;
            --color-blue-50: #eff6ff !important;
            --color-blue-100: #dbeafe !important;
            --color-blue-600: #2563eb !important;
            --color-blue-950: #172554 !important;
          }
          .print-container * {
            box-sizing: border-box;
            border-color: #171717; /* Default to solid dark for borders */
          }
          
          /* Forced HEX overrides for PDF rendering engine */
          .bg-neutral-900 { background-color: #171717 !important; }
          .text-neutral-900 { color: #171717 !important; }
          .text-neutral-800 { color: #262626 !important; }
          .text-neutral-500 { color: #737373 !important; }
          .text-neutral-400 { color: #a3a3a3 !important; }
          .bg-neutral-100 { background-color: #f5f5f5 !important; }
          .bg-neutral-50 { background-color: #fafafa !important; }
          .border-neutral-900 { border-color: #171717 !important; }
          .border-neutral-300 { border-color: #d4d4d4 !important; }
          .border-neutral-200 { border-color: #e5e5e5 !important; }
          .border-neutral-100 { border-color: #f5f5f5 !important; }
          .text-blue-950 { color: #020617 !important; }
          .text-white { color: #ffffff !important; }
        `}</style>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-neutral-50 text-neutral-900 font-sans p-6 md:p-12">
      <div className="max-w-6xl mx-auto">
        <header className="flex flex-col md:flex-row md:items-center justify-between gap-6 mb-12">
          <div>
            <h1 className="text-4xl font-bold tracking-tight text-neutral-900">Delivery Label Master</h1>
            <p className="text-neutral-500 mt-2">Upload Excel lists to generate professional A4 delivery labels.</p>
          </div>
          
          <div className="flex flex-wrap items-center gap-3">
            <input 
              type="file" 
              accept=".xlsx, .xls" 
              onChange={handleFileUpload} 
              className="hidden" 
              ref={fileInputRef}
            />
             <input 
              type="file" 
              accept=".xlsx, .xls" 
              onChange={handleFileUpload} 
              className="hidden" 
              ref={contactInputRef}
            />
            
            <div className="flex gap-2">
              <button 
                onClick={() => setIsPrinting(true)}
                disabled={selectedIds.size === 0}
                className="flex items-center gap-3 px-8 py-3 bg-neutral-900 text-white rounded-xl hover:bg-neutral-800 shadow-xl shadow-neutral-900/10 disabled:opacity-30 disabled:cursor-not-allowed transition-all font-bold text-sm"
              >
                <Printer size={20} className="animate-pulse" />
                Generate PDF & Print
              </button>
            </div>
          </div>
        </header>

        <main className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Data Import Panel */}
          <section className="lg:col-span-1 space-y-6">
            <div 
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              className={`relative bg-white p-6 rounded-2xl border-2 border-dashed transition-all group ${
                isDragging ? 'border-blue-500 bg-blue-50 shadow-2xl scale-[1.02]' : 'border-neutral-200 hover:border-neutral-400'
              }`}
            >
              <h2 className="text-lg font-bold mb-4 flex items-center gap-2">
                <Upload className="text-neutral-900" size={20} />
                Import Core
              </h2>
              
              <div className="space-y-4">
                <div className="grid grid-cols-2 gap-2">
                  <button 
                    onClick={() => fileInputRef.current?.click()}
                    className="flex flex-col items-center justify-center gap-2 p-4 bg-blue-50 border border-blue-100 rounded-xl hover:bg-blue-100 transition-colors group/btn"
                  >
                    <FileSpreadsheet className="text-blue-600 group-hover/btn:scale-110 transition-transform" size={24} />
                    <span className="text-[10px] font-black uppercase text-blue-700">Add Samples</span>
                  </button>
                  <button 
                    onClick={() => contactInputRef.current?.click()}
                    className="flex flex-col items-center justify-center gap-2 p-4 bg-green-50 border border-green-100 rounded-xl hover:bg-green-100 transition-colors group/btn"
                  >
                    <Upload className="text-green-600 group-hover/btn:scale-110 transition-transform" size={24} />
                    <span className="text-[10px] font-black uppercase text-green-700">Add Contacts</span>
                  </button>
                </div>

                <div className="relative">
                  <textarea 
                    onPaste={handlePaste}
                    placeholder="Paste Samples or Contacts from Excel here..."
                    className="w-full h-32 p-4 text-xs font-mono bg-neutral-50 border border-neutral-100 rounded-xl focus:ring-2 focus:ring-neutral-900 focus:bg-white focus:outline-none transition-all resize-none"
                  ></textarea>
                  <div className="absolute top-3 right-3 pointer-events-none opacity-50">
                    <CheckCircle2 size={16} className="text-neutral-300" />
                  </div>
                </div>

                <div className="text-[10px] text-center text-neutral-400 font-medium">
                  {isDragging ? 'RELEASE TO UPLOAD' : 'DRAG EXCEL FILES HERE OR PASTE ROWS'}
                </div>

                <div className="pt-4 border-t border-neutral-100">
                  <h3 className="text-[10px] font-black uppercase text-neutral-400 mb-2 tracking-widest">Manual Entry</h3>
                  <div className="space-y-4">
                    {/* Manual Sample */}
                    <div className="bg-neutral-50 p-3 rounded-xl border border-neutral-100 space-y-2">
                      <div className="text-[9px] font-bold text-blue-600 uppercase">Single Sample / Smart Paste</div>
                      <div className="grid grid-cols-2 gap-2">
                        <input 
                          type="text" 
                          placeholder="Sample PO"
                          value={manualSample.samplePo}
                          onPaste={handleManualSamplePaste}
                          onChange={e => setManualSample({...manualSample, samplePo: e.target.value})}
                          className="text-[11px] p-2 bg-white border border-neutral-200 rounded-lg focus:outline-none"
                        />
                        <input 
                          type="text" 
                          placeholder="PI No"
                          value={manualSample.piNo}
                          onPaste={handleManualSamplePaste}
                          onChange={e => setManualSample({...manualSample, piNo: e.target.value})}
                          className="text-[11px] p-2 bg-white border border-neutral-200 rounded-lg focus:outline-none"
                        />
                        <input 
                          type="text" 
                          placeholder="Customer"
                          value={manualSample.customer}
                          onPaste={handleManualSamplePaste}
                          onChange={e => setManualSample({...manualSample, customer: e.target.value})}
                          className="text-[11px] p-2 bg-white border border-neutral-200 rounded-lg focus:outline-none col-span-2"
                        />
                        <input 
                          type="text" 
                          placeholder="Type (e.g. PRICE BARCODE)"
                          value={manualSample.sampleType}
                          onPaste={handleManualSamplePaste}
                          onChange={e => setManualSample({...manualSample, sampleType: e.target.value})}
                          className="text-[11px] p-2 bg-white border border-neutral-200 rounded-lg focus:outline-none col-span-2"
                        />
                      </div>
                      <button 
                        onClick={addManualSample}
                        className="w-full bg-blue-600 text-white text-[10px] font-bold uppercase py-2 rounded-lg hover:bg-blue-700 transition-colors shadow-lg shadow-blue-600/20"
                      >
                        Add to List
                      </button>
                    </div>

                    {/* Manual Contact */}
                    <div className="bg-neutral-50 p-3 rounded-xl border border-neutral-100 space-y-2">
                      <div className="text-[9px] font-bold text-green-600 uppercase">Single Contact / Smart Paste</div>
                      <div className="grid grid-cols-1 gap-2">
                        <input 
                          type="text" 
                          placeholder="Customer Name"
                          value={manualContact.customerName}
                          onPaste={handleManualContactPaste}
                          onChange={e => setManualContact({...manualContact, customerName: e.target.value})}
                          className="text-[11px] p-2 bg-white border border-neutral-200 rounded-lg focus:outline-none"
                        />
                        <input 
                          type="text" 
                          placeholder="Contact Person (Attn)"
                          value={manualContact.contactPerson}
                          onPaste={handleManualContactPaste}
                          onChange={e => setManualContact({...manualContact, contactPerson: e.target.value})}
                          className="text-[11px] p-2 bg-white border border-neutral-200 rounded-lg focus:outline-none"
                        />
                        <input 
                          type="text" 
                          placeholder="Phone Number"
                          value={manualContact.phoneNumber}
                          onPaste={handleManualContactPaste}
                          onChange={e => setManualContact({...manualContact, phoneNumber: e.target.value})}
                          className="text-[11px] p-2 bg-white border border-neutral-200 rounded-lg focus:outline-none"
                        />
                        <input 
                          type="text" 
                          placeholder="Address (Optional)"
                          value={manualContact.address}
                          onPaste={handleManualContactPaste}
                          onChange={e => setManualContact({...manualContact, address: e.target.value})}
                          className="text-[11px] p-2 bg-white border border-neutral-200 rounded-lg focus:outline-none"
                        />
                      </div>
                      <button 
                        onClick={addManualContact}
                        className="w-full bg-green-600 text-white text-[10px] font-bold uppercase py-2 rounded-lg hover:bg-green-700 transition-colors shadow-lg shadow-green-600/20"
                      >
                        Save Contact
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            <div className="bg-white p-6 rounded-2xl border border-neutral-200 shadow-sm">
              <h2 className="text-xs font-black uppercase text-neutral-400 mb-6 tracking-widest flex items-center justify-between">
                System Status
                <span className="w-2 h-2 rounded-full bg-green-500 animate-pulse"></span>
              </h2>
              
              <div className="space-y-4">
                <div className="flex justify-between items-center p-3 bg-neutral-50 rounded-xl border border-neutral-100">
                  <div className="flex flex-col">
                    <span className="text-neutral-500 text-xs font-bold uppercase">Samples</span>
                    <span className="font-bold text-lg">{samples.length}</span>
                  </div>
                  <button 
                    onClick={() => downloadTemplate('samples')}
                    className="text-[10px] bg-blue-50 text-blue-700 px-2 py-1 rounded hover:bg-blue-100 font-bold uppercase transition-colors"
                  >
                    Template
                  </button>
                </div>

                <div className="flex justify-between items-center p-3 bg-neutral-50 rounded-xl border border-neutral-100">
                  <div className="flex flex-col">
                    <span className="text-neutral-500 text-xs font-bold uppercase">Contacts</span>
                    <span className="font-bold text-lg">{contacts.length}</span>
                  </div>
                  <div className="flex gap-2">
                    <button 
                      onClick={() => downloadTemplate('contacts')}
                      className="text-[10px] bg-green-50 text-green-700 px-2 py-1 rounded hover:bg-green-100 font-bold uppercase transition-colors"
                    >
                      Template
                    </button>
                    {contacts.length > 0 && (
                      <button 
                        onClick={() => updateContacts([])}
                        className="text-[10px] bg-red-50 text-red-600 px-2 py-1 rounded hover:bg-red-100 font-bold uppercase transition-colors"
                        title="Delete all saved contacts"
                      >
                        Reset
                      </button>
                    )}
                  </div>
                </div>

                <div className="flex justify-between items-center p-4 bg-neutral-900 text-white rounded-xl shadow-lg ring-4 ring-neutral-900/5">
                  <span className="text-neutral-400 text-sm font-medium">Ready to Print</span>
                  <span className="text-2xl font-black">{selectedIds.size}</span>
                </div>
              </div>
              
              <div className="mt-8 pt-6 border-t border-neutral-100 space-y-4">
                <div className="space-y-2">
                  <h3 className="text-xs font-bold text-neutral-400 uppercase tracking-widest">How to use</h3>
                  <ol className="text-xs text-neutral-500 space-y-2 leading-relaxed list-decimal pl-4">
                    <li>Download the <strong>Contact List Template</strong> and fill with your customer info.</li>
                    <li>Upload the saved Contacts file. (Saved automatically for next time!)</li>
                    <li>Download <strong>Sample PO Template</strong> and fill with order data.</li>
                    <li>Upload Sample POs, select labels, and click <strong>Print</strong>.</li>
                  </ol>
                </div>
                
                <button 
                  onClick={clearData}
                  className="w-full text-red-500 hover:text-red-700 text-xs flex items-center justify-center gap-2 p-3 border border-dashed border-neutral-200 hover:border-red-200 rounded-xl transition-all"
                >
                  <Trash2 size={14} /> Clear Current Working List
                </button>
              </div>
            </div>

            {contacts.length > 0 && (
              <div className="bg-white p-6 rounded-2xl border border-neutral-200 shadow-sm flex flex-col max-h-[500px]">
                <h2 className="text-sm font-bold uppercase text-neutral-400 mb-4 tracking-tighter">Saved Customers ({contacts.length})</h2>
                <div className="overflow-y-auto space-y-2 pr-2 custom-scrollbar">
                  {contacts.map((contact, idx) => (
                    <div key={idx} className="group relative">
                      {editingContactIndex === idx ? (
                        <div className="p-3 bg-blue-50/50 rounded-xl border border-blue-100 space-y-2">
                          <input 
                            value={editContactBuffer?.customerName || ''} 
                            onChange={e => setEditContactBuffer(prev => prev ? {...prev, customerName: e.target.value} : null)}
                            className="w-full text-[11px] font-bold p-2 bg-white border border-blue-200 rounded-lg outline-none"
                            placeholder="Customer Name"
                          />
                          <input 
                            value={editContactBuffer?.contactPerson || ''} 
                            onChange={e => setEditContactBuffer(prev => prev ? {...prev, contactPerson: e.target.value} : null)}
                            className="w-full text-[10px] p-2 bg-white border border-blue-200 rounded-lg outline-none"
                            placeholder="Contact Person"
                          />
                          <input 
                            value={editContactBuffer?.phoneNumber || ''} 
                            onChange={e => setEditContactBuffer(prev => prev ? {...prev, phoneNumber: e.target.value} : null)}
                            className="w-full text-[10px] p-2 bg-white border border-blue-200 rounded-lg outline-none"
                            placeholder="Phone Number"
                          />
                          <input 
                            value={editContactBuffer?.address || ''} 
                            onChange={e => setEditContactBuffer(prev => prev ? {...prev, address: e.target.value} : null)}
                            className="w-full text-[10px] p-2 bg-white border border-blue-200 rounded-lg outline-none"
                            placeholder="Address"
                          />
                          <div className="flex gap-2 pt-1">
                            <button 
                              onClick={saveEditedContact}
                              className="flex-1 bg-blue-600 text-white text-[9px] font-bold uppercase py-1.5 rounded-lg hover:bg-blue-700 transition-colors"
                            >
                              Save
                            </button>
                            <button 
                              onClick={cancelEditingContact}
                              className="bg-neutral-200 text-neutral-600 text-[9px] font-bold uppercase px-3 py-1.5 rounded-lg hover:bg-neutral-300 transition-colors"
                            >
                              Cancel
                            </button>
                          </div>
                        </div>
                      ) : (
                        <div 
                          onClick={() => startEditingContact(idx)}
                          className="p-3 bg-neutral-50 rounded-xl border border-neutral-100 group-hover:border-neutral-200 group-hover:bg-neutral-100/50 transition-all cursor-pointer relative"
                        >
                          <div className="text-[11px] font-black text-neutral-900 leading-tight pr-6">{contact.customerName}</div>
                          <div className="text-[9px] font-bold text-neutral-400 truncate">
                            {contact.contactPerson} | <span className="opacity-50 font-medium italic">{contact.phoneNumber}</span>
                          </div>
                          {contact.address && (
                            <div className="text-[8px] text-neutral-400/70 truncate pt-0.5 italic">{contact.address}</div>
                          )}
                          <button 
                            onClick={(e) => { e.stopPropagation(); deleteContact(idx); }}
                            className="absolute right-2 top-2 w-6 h-6 items-center justify-center bg-red-50 text-red-500 rounded-full opacity-0 group-hover:opacity-100 flex transition-all hover:bg-red-500 hover:text-white"
                          >
                            <Trash2 size={12} />
                          </button>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            )}
          </section>

          {/* List Section */}
          <section className="lg:col-span-2 space-y-4">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" size={18} />
              <input 
                type="text" 
                placeholder="Search by customer or PO..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-10 pr-4 py-3 bg-white border border-neutral-200 rounded-xl focus:ring-2 focus:ring-neutral-900 focus:outline-none transition-shadow"
              />
            </div>

            <div className="bg-white rounded-2xl border border-neutral-200 shadow-sm overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-neutral-50 border-b border-neutral-100">
                      <th className="px-6 py-4">
                        <input 
                          type="checkbox" 
                          checked={selectedIds.size === filteredSamples.length && filteredSamples.length > 0}
                          onChange={selectAll}
                          className="w-4 h-4 accent-neutral-900"
                        />
                      </th>
                      <th className="px-4 py-4 text-xs font-bold text-neutral-400 uppercase tracking-wider">Info</th>
                      <th className="px-4 py-4 text-xs font-bold text-neutral-400 uppercase tracking-wider">Customer</th>
                      <th className="px-4 py-4 text-xs font-bold text-neutral-400 uppercase tracking-wider text-right">Details</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-neutral-50">
                    {filteredSamples.length === 0 ? (
                      <tr>
                        <td colSpan={4} className="px-6 py-12 text-center text-neutral-400 italic">
                          No samples data. Please upload an Excel file.
                        </td>
                      </tr>
                    ) : (
                      filteredSamples.map((sample) => {
                        const isSelected = selectedIds.has(sample.id);
                        const isEditing = editingSampleId === sample.id;
                        const hasContact = contacts.some(c => c.customerName.trim().toLowerCase() === sample.customer.trim().toLowerCase());

                        if (isEditing) {
                          return (
                            <tr key={sample.id} className="bg-blue-50/50">
                              <td className="px-6 py-4">
                                <div className="flex items-center gap-2">
                                  <button 
                                    onClick={(e) => { e.stopPropagation(); saveEditedSample(); }}
                                    className="p-1 bg-green-500 text-white rounded hover:bg-green-600 shadow-sm"
                                  >
                                    <CheckCircle2 size={14} />
                                  </button>
                                  <button 
                                    onClick={(e) => { e.stopPropagation(); cancelEditingSample(); }}
                                    className="p-1 bg-neutral-400 text-white rounded hover:bg-neutral-500 shadow-sm"
                                  >
                                    <X size={14} />
                                  </button>
                                </div>
                              </td>
                              <td className="px-4 py-4 space-y-1">
                                <input 
                                  value={editSampleBuffer?.samplePo || ''} 
                                  onChange={e => setEditSampleBuffer(prev => prev ? {...prev, samplePo: e.target.value} : null)}
                                  onClick={e => e.stopPropagation()}
                                  className="w-full text-xs font-mono p-1 border border-blue-200 rounded outline-none bg-white"
                                  placeholder="Sample PO"
                                />
                                <input 
                                  value={editSampleBuffer?.piNo || ''} 
                                  onChange={e => setEditSampleBuffer(prev => prev ? {...prev, piNo: e.target.value} : null)}
                                  onClick={e => e.stopPropagation()}
                                  className="w-full text-[10px] p-1 border border-blue-200 rounded outline-none bg-white"
                                  placeholder="PI No"
                                />
                              </td>
                              <td className="px-4 py-4">
                                <input 
                                  value={editSampleBuffer?.customer || ''} 
                                  onChange={e => setEditSampleBuffer(prev => prev ? {...prev, customer: e.target.value} : null)}
                                  onClick={e => e.stopPropagation()}
                                  className="w-full text-xs p-1 border border-blue-200 rounded outline-none bg-white font-bold"
                                  placeholder="Customer"
                                />
                              </td>
                              <td className="px-4 py-4 text-right">
                                <input 
                                  value={editSampleBuffer?.sampleType || ''} 
                                  onChange={e => setEditSampleBuffer(prev => prev ? {...prev, sampleType: e.target.value} : null)}
                                  onClick={e => e.stopPropagation()}
                                  className="w-24 text-[10px] p-1 border border-blue-200 rounded outline-none bg-white text-right"
                                  placeholder="Type"
                                />
                              </td>
                            </tr>
                          );
                        }

                        return (
                          <tr 
                            key={sample.id} 
                            onClick={() => toggleSelect(sample.id)}
                            className={`hover:bg-neutral-50 cursor-pointer transition-colors group ${isSelected ? 'bg-blue-50/20' : ''}`}
                          >
                            <td className="px-6 py-4">
                              <div className="flex items-center gap-3">
                                <input 
                                  type="checkbox" 
                                  checked={isSelected}
                                  readOnly
                                  className="w-4 h-4 accent-neutral-900"
                                />
                                <button 
                                  onClick={(e) => { e.stopPropagation(); startEditingSample(sample); }}
                                  className="opacity-0 group-hover:opacity-100 p-1 text-neutral-400 hover:text-blue-600 transition-opacity"
                                  title="Edit Sample"
                                >
                                  <Edit3 size={14} />
                                </button>
                              </div>
                            </td>
                            <td className="px-4 py-4">
                              <div className="font-mono text-sm font-semibold">{sample.samplePo}</div>
                              <div className="text-xs text-neutral-400">{sample.piNo}</div>
                            </td>
                            <td className="px-4 py-4">
                              <div className="font-medium text-neutral-900">{sample.customer}</div>
                              {hasContact ? (
                                <span className="inline-flex items-center gap-1 text-[10px] bg-green-100 text-green-700 px-1.5 py-0.5 rounded uppercase font-bold">
                                  Matched
                                </span>
                              ) : (
                                <span className="inline-flex items-center gap-1 text-[10px] bg-amber-100 text-amber-700 px-1.5 py-0.5 rounded uppercase font-bold">
                                  No Contact Info
                                </span>
                              )}
                            </td>
                            <td className="px-4 py-4 text-right">
                              <div className="flex items-center justify-end gap-2">
                                <div className="text-xs text-neutral-500">{sample.sampleType}</div>
                                <button 
                                  onClick={(e) => { e.stopPropagation(); deleteSample(sample.id); }}
                                  className="opacity-0 group-hover:opacity-100 p-1 text-neutral-400 hover:text-red-500 transition-opacity"
                                  title="Delete Row"
                                >
                                  <Trash2 size={14} />
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        </main>
      </div>
    </div>
  );
}
