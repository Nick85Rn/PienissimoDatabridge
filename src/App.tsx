import React, { useState, type ChangeEvent } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { 
  Upload, Download, CheckCircle2, AlertCircle, FileSpreadsheet, ArrowRight, 
  Layout, Loader2, X, Sparkles, ShieldCheck, MapPin, User, Tag, RefreshCw,
  Wand2, Settings2, Split, CopyMinus
} from 'lucide-react';

// --- CONFIGURAZIONE ---
interface DemoColumn {
  id: string; 
  label: string; 
  category: 'Anagrafica' | 'Contatti' | 'Indirizzo' | 'Dettagli';
  requiredGroup?: string; 
  description?: string;
  cleanType?: 'text' | 'phone' | 'email' | 'none';
}

const DEMO_COLUMNS: DemoColumn[] = [
  // ANAGRAFICA
  { id: 'nome', label: 'nome', category: 'Anagrafica', description: 'Nome (o Nome Completo se attivi Split)', cleanType: 'text' },
  { id: 'cognome', label: 'cognome', category: 'Anagrafica', description: 'Cognome (lascia vuoto se usi Split)', cleanType: 'text' },
  { id: 'data_nascita', label: 'data di nascita', category: 'Anagrafica', description: 'Formato: GG/MM/AAAA', cleanType: 'none' },

  // CONTATTI
  { id: 'email', label: 'email', category: 'Contatti', requiredGroup: 'contatto', description: 'Email personale/aziendale', cleanType: 'email' },
  { id: 'telefono', label: 'telefono', category: 'Contatti', requiredGroup: 'contatto', description: 'Cellulare o fisso', cleanType: 'phone' },

  // INDIRIZZO
  { id: 'indirizzo', label: 'indirizzo', category: 'Indirizzo', description: 'Via e numero civico', cleanType: 'text' },
  { id: 'citta', label: 'citta', category: 'Indirizzo', description: 'Città di residenza', cleanType: 'text' },
  { id: 'cap', label: 'cap', category: 'Indirizzo', description: 'Codice Postale', cleanType: 'none' },
  { id: 'provincia', label: 'provincia', category: 'Indirizzo', description: 'Provincia (es. RN)', cleanType: 'text' },
  { id: 'nazione', label: 'nazione', category: 'Indirizzo', description: 'Nazione (es. Italia)', cleanType: 'text' },

  // DETTAGLI
  { id: 'note', label: 'note', category: 'Dettagli', description: 'Note libere aggiuntive', cleanType: 'none' },
  { id: 'profilazione', label: 'profilazione', category: 'Dettagli', description: 'Campo custom per segmentazione', cleanType: 'none' },
  { id: 'tags', label: 'tags', category: 'Dettagli', description: 'Etichette separate da virgola', cleanType: 'none' },
];

type Mapping = Record<string, string>;

export default function App() {
  const [step, setStep] = useState<1 | 2>(1); 
  const [fileData, setFileData] = useState<any[]>([]);
  const [userHeaders, setUserHeaders] = useState<string[]>([]);
  const [headerSamples, setHeaderSamples] = useState<Record<string, string>>({}); 
  const [mapping, setMapping] = useState<Mapping>({});
  
  // --- STATI AVANZATI ---
  const [cleaningEnabled, setCleaningEnabled] = useState<boolean>(true);
  const [splitNameEnabled, setSplitNameEnabled] = useState<boolean>(false);
  const [deduplicateEnabled, setDeduplicateEnabled] = useState<boolean>(false);
  const [globalTag, setGlobalTag] = useState<string>("");
  const [defaultPrefix, setDefaultPrefix] = useState<string>("+39");
  
  const [previewRows, setPreviewRows] = useState<any[]>([]);
  const [originalFileName, setOriginalFileName] = useState<string>("");
  const [customFileName, setCustomFileName] = useState<string>("");
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showSuccessToast, setShowSuccessToast] = useState(false);

  // --- MOTORE DI PULIZIA DATI ---
  const cleanValue = (value: any, type: string | undefined) => {
    if (value === null || value === undefined) return "";
    const str = String(value).trim();
    if (str === "") return "";

    if (!cleaningEnabled) return str;

    switch (type) {
      case 'text': return str.toLowerCase().replace(/(?:^|\s|['-])\S/g, (a) => a.toUpperCase());
      case 'email': return str.toLowerCase();
      case 'phone': {
        let cleanPhone = str.replace(/[^0-9+]/g, '');
        if (cleanPhone.startsWith('00')) cleanPhone = '+' + cleanPhone.substring(2);
        if (cleanPhone.startsWith('+')) return cleanPhone;
        if (cleanPhone.length > 5) return defaultPrefix + cleanPhone;
        return cleanPhone;
      }
      default: return str;
    }
  };

  // --- LOGICA CORE ELABORAZIONE RIGA ---
  const processRow = (row: any, currentMapping: Mapping) => {
    const newRow: any = {};
    
    // Gestione Split Nome/Cognome
    const splitParts = { nome: "", cognome: "" };
    if (splitNameEnabled && currentMapping['nome'] && !currentMapping['cognome']) {
      const fullName = row[currentMapping['nome']] || "";
      const parts = String(fullName).trim().split(' ');
      splitParts.nome = parts[0] || "";
      splitParts.cognome = parts.slice(1).join(' ') || "";
    }

    DEMO_COLUMNS.forEach(col => {
      let rawValue = "";

      if (splitNameEnabled && col.id === 'nome' && splitParts.nome) {
        rawValue = splitParts.nome;
      } else if (splitNameEnabled && col.id === 'cognome' && splitParts.cognome) {
        rawValue = splitParts.cognome;
      } else {
        const mappedHeader = currentMapping[col.id];
        rawValue = mappedHeader ? row[mappedHeader] : "";
      }

      // Gestione Global Tag
      if (col.id === 'tags' && globalTag) {
        const existingTags = rawValue ? rawValue + ", " : "";
        rawValue = existingTags + globalTag;
      }

      newRow[col.label] = cleanValue(rawValue, col.cleanType);
    });

    return newRow;
  };

  const processAllData = (dataSource: any[], currentMapping: Mapping) => {
    let processed = dataSource.map(row => processRow(row, currentMapping));

    // Gestione Deduplica
    if (deduplicateEnabled) {
      const uniqueMap = new Map();
      processed.forEach(row => {
        const key = row['email'] || row['telefono']; 
        if (key && key.length > 5) {
          if (!uniqueMap.has(key)) uniqueMap.set(key, row);
        } else {
           uniqueMap.set(Math.random(), row); 
        }
      });
      processed = Array.from(uniqueMap.values());
    }

    return processed;
  };

  // --- LOGICA DI IMPORT ---
  const processFile = (file: File) => {
    setIsProcessing(true);
    setOriginalFileName(file.name);
    setCustomFileName(file.name.split('.')[0]); 
    const extension = file.name.split('.').pop()?.toLowerCase();
    
    const parseCallback = (data: any[], headers: string[]) => {
      setTimeout(() => {
        setUserHeaders(headers); 
        setFileData(data); 
        
        const firstRow = data.find(row => Object.values(row).some(v => v));
        const samples: Record<string, string> = {};
        if (firstRow) {
          headers.forEach(h => {
            const val = firstRow[h];
            samples[h] = val ? String(val).substring(0, 15) + (String(val).length > 15 ? '...' : '') : '-';
          });
        }
        setHeaderSamples(samples);

        const initialMapping = generateAutoMatch(headers);
        setMapping(initialMapping);
        
        // Genera anteprima iniziale
        const initialProcessed = processAllData(data, initialMapping);
        setPreviewRows(initialProcessed.slice(0, 30));

        setStep(2); 
        setIsProcessing(false);
      }, 800);
    };

    if (extension === 'csv') {
      // FIX: Tipizzare 'results' come 'any' per evitare errore TS7006
      Papa.parse(file, { 
        header: true, 
        skipEmptyLines: true, 
        complete: (results: any) => parseCallback(results.data, results.meta.fields || []) 
      });
    } else if (['xlsx', 'xls'].includes(extension || '')) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        const wb = XLSX.read(evt.target?.result, { type: 'binary' });
        const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        const headers = data.length > 0 ? Object.keys(data[0] as object) : [];
        parseCallback(data, headers);
      };
      reader.readAsBinaryString(file);
    } else { setIsProcessing(false); alert("Formato non supportato."); }
  };

  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>) => { const file = e.target.files?.[0]; if (file) processFile(file); };
  const onDrag = (e: React.DragEvent, dragging: boolean) => { e.preventDefault(); setIsDragging(dragging); };
  const onDrop = (e: React.DragEvent) => { e.preventDefault(); setIsDragging(false); const file = e.dataTransfer.files?.[0]; if (file) processFile(file); };

  const generateAutoMatch = (headers: string[]) => {
    const newMapping: Mapping = {};
    DEMO_COLUMNS.forEach(col => {
      const match = headers.find(h => {
        const hLow = h.toLowerCase().trim();
        const idLow = col.id.toLowerCase();
        if (idLow === 'data_nascita' && (hLow.includes('nascita') || hLow.includes('birth'))) return true;
        if (idLow === 'cap' && (hLow.includes('zip') || hLow.includes('postale'))) return true;
        if (idLow === 'indirizzo' && (hLow.includes('via') || hLow.includes('address'))) return true;
        return hLow === col.label.toLowerCase() || hLow === idLow || 
          (col.id === 'email' && (hLow.includes('mail') || hLow.includes('e-mail'))) ||
          (col.id === 'telefono' && (hLow.includes('tel') || hLow.includes('cel') || hLow.includes('phone')));
      });
      if (match) newMapping[col.id] = match;
    });
    return newMapping;
  };

  const isMappingValid = () => !!(mapping['email'] || mapping['telefono']);

  const handleManualUpdate = () => {
    const fullProcessed = processAllData(fileData, mapping);
    setPreviewRows(fullProcessed.slice(0, 30));
  };

  // --- ESPORTAZIONE ---
  const exportFile = () => {
    const fullProcessed = processAllData(fileData, mapping);
    
    const now = new Date();
    const ts = `${now.getFullYear()}${(now.getMonth()+1).toString().padStart(2,'0')}${now.getDate().toString().padStart(2,'0')}_${now.getHours()}${now.getMinutes()}`;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([Papa.unparse(fullProcessed)], { type: 'text/csv;charset=utf-8;' }));
    link.download = `${customFileName}_${ts}.csv`;
    link.click();
    setShowSuccessToast(true); setTimeout(() => setShowSuccessToast(false), 4000);
  };
  
  const getCategoryIcon = (cat: string) => {
    switch(cat) {
      case 'Anagrafica': return <User size={14} />;
      case 'Contatti': return <ShieldCheck size={14} />;
      case 'Indirizzo': return <MapPin size={14} />;
      case 'Dettagli': return <Tag size={14} />;
      default: return <Layout size={14} />;
    }
  };

  const getCategoryColor = (cat: string) => {
    switch(cat) {
      case 'Anagrafica': return 'bg-blue-100 text-blue-700 border-blue-200';
      case 'Contatti': return 'bg-emerald-100 text-emerald-700 border-emerald-200';
      case 'Indirizzo': return 'bg-amber-100 text-amber-700 border-amber-200';
      case 'Dettagli': return 'bg-purple-100 text-purple-700 border-purple-200';
      default: return 'bg-slate-100 text-slate-700';
    }
  };
  
  return (
    <div className="min-h-screen bg-slate-100 text-slate-900 font-sans pb-40 flex flex-col justify-between">
      {showSuccessToast && (
        <div className="fixed top-6 right-6 z-[100] animate-in slide-in-from-top-4 fade-in duration-300">
          <div className="bg-white border-l-4 border-green-600 shadow-2xl rounded-lg p-4 flex items-start gap-3 max-w-sm">
            <div className="bg-green-100 p-1.5 rounded-full text-green-700"><Sparkles size={18} /></div>
            <div><h4 className="font-bold text-lg">Download Completato!</h4><p className="text-sm text-slate-600 mt-1">File pronto per l'importazione.</p></div>
            <button onClick={() => setShowSuccessToast(false)} className="text-slate-400 hover:text-slate-700 ml-2"><X size={18} /></button>
          </div>
        </div>
      )}

      <div className="max-w-[95%] xl:max-w-7xl mx-auto px-4 w-full py-8">
        <header className="flex flex-col md:flex-row items-center justify-between py-6 px-8 mb-8 bg-slate-900 rounded-3xl shadow-xl">
            <div className="flex items-center gap-3 mb-4 md:mb-0">
                <div className="p-2 bg-white/10 rounded-full backdrop-blur-sm"><ShieldCheck className="text-emerald-400" size={32} /></div>
                <div><h1 className="text-2xl font-black text-white tracking-wide leading-none">Pienissimo</h1><span className="text-xs font-bold text-emerald-400 uppercase tracking-[0.2em]">DataBridge</span></div>
            </div>
            <div className="text-center md:text-right">
                <h2 className="text-slate-300 font-medium text-sm mb-1">Normalizzatore Contatti Enterprise</h2>
                <div className="flex items-center gap-2 justify-center md:justify-end rounded-full bg-slate-800/50 py-1 px-3 inline-flex">
                     <span className="h-2 w-2 rounded-full bg-emerald-500 animate-pulse shadow-[0_0_8px_rgba(16,185,129,0.6)]"></span><span className="text-xs text-emerald-400 font-bold uppercase">Sistema Attivo</span>
                </div>
            </div>
        </header>

        {isProcessing && (
          <div className="fixed inset-0 bg-slate-950/70 backdrop-blur-md z-50 flex flex-col items-center justify-center text-white">
            <Loader2 className="animate-spin text-emerald-500 mb-4" size={64} /><p className="text-2xl font-black">Elaborazione...</p><p className="text-slate-300 mt-2 font-medium">Analisi struttura file in corso</p>
          </div>
        )}

        <div className="bg-white rounded-3xl shadow-2xl border-2 border-slate-200 overflow-hidden min-h-[600px]">
            {step === 1 && (
            <div className="h-full flex flex-col items-center justify-center p-8 md:p-16 animate-in fade-in zoom-in-95 duration-500 bg-slate-50">
                <div className="text-center mb-10">
                    <h3 className="text-4xl font-black text-slate-900 mb-4">Importa Database Clienti</h3>
                    <p className="text-slate-700 max-w-lg mx-auto text-lg font-medium leading-relaxed">Carica il file Excel o CSV. Verrà convertito nel formato standard Pienissimo.</p>
                </div>
                <label onDragOver={(e) => onDrag(e, true)} onDragLeave={(e) => onDrag(e, false)} onDrop={onDrop}
                className={`group relative flex flex-col items-center justify-center w-full max-w-2xl h-72 rounded-3xl cursor-pointer transition-all duration-300 ${isDragging ? 'bg-emerald-50 border-4 border-emerald-500 scale-[1.02] shadow-2xl' : 'bg-white border-4 border-dashed border-slate-300 hover:border-emerald-500 hover:bg-emerald-50/30 hover:shadow-lg'}`}>
                <div className={`p-5 rounded-full mb-5 transition-colors duration-300 shadow-sm ${isDragging ? 'bg-emerald-200 text-emerald-800' : 'bg-slate-100 border-2 border-slate-200 text-slate-700 group-hover:text-emerald-700 group-hover:border-emerald-400 group-hover:scale-110'}`}><Upload size={40} strokeWidth={2.5} /></div>
                <h4 className="font-black text-2xl text-slate-800 group-hover:text-emerald-800 transition-colors">{isDragging ? 'Rilascia il file ora' : 'Clicca o trascina qui'}</h4>
                <p className="text-sm text-slate-700 mt-4 font-bold bg-slate-200 px-5 py-2 rounded-full border-2 border-slate-300 group-hover:border-emerald-300 transition-colors">CSV, XLS, XLSX supportati</p>
                <input type="file" className="hidden" accept=".csv, .xls, .xlsx" onChange={handleFileUpload} />
                </label>
            </div>
            )}

            {step === 2 && (
            <div className="animate-in slide-in-from-right-8 duration-500 flex flex-col h-full bg-slate-50">
                
                {/* TOOLBAR AVANZATA */}
                <div className="bg-white border-b-2 border-slate-200 p-5 md:px-6 flex flex-col xl:flex-row items-start xl:items-center justify-between gap-5 shadow-sm z-10 sticky top-0">
                    
                    <div className="flex flex-wrap gap-4 w-full">
                         {/* FILE INFO */}
                        <div className="flex items-center gap-3 bg-slate-100 p-2.5 rounded-xl border border-slate-200">
                             <div className="bg-white p-2 rounded-lg border border-slate-300 text-emerald-700"><FileSpreadsheet size={20}/></div>
                             <div className="max-w-[150px]"><p className="text-[10px] font-black text-slate-500 uppercase">File</p><p className="font-bold text-slate-900 text-xs truncate">{originalFileName}</p></div>
                        </div>

                        {/* PANEL: PULIZIA */}
                        <div className="flex-1 min-w-[280px] bg-slate-50 border-2 border-slate-200 rounded-xl px-4 py-2 flex flex-col justify-center">
                            <div className="flex items-center gap-2 mb-1">
                                <Wand2 size={14} className="text-indigo-600" />
                                <span className="text-[10px] font-black text-slate-500 uppercase">Pulizia Dati</span>
                            </div>
                            <div className="flex items-center gap-3">
                                <label className="flex items-center gap-1.5 cursor-pointer text-xs font-bold text-slate-700 hover:text-indigo-700">
                                    <input type="checkbox" checked={cleaningEnabled} onChange={(e) => setCleaningEnabled(e.target.checked)} className="rounded text-indigo-600 focus:ring-indigo-500" /> Auto-Fix
                                </label>
                                {cleaningEnabled && (
                                    <input type="text" value={defaultPrefix} onChange={(e) => setDefaultPrefix(e.target.value)} className="w-16 px-1.5 py-0.5 text-xs font-bold border border-slate-300 rounded text-center" placeholder="+39" title="Prefisso Default" />
                                )}
                            </div>
                        </div>

                        {/* PANEL: TOOLS AVANZATI (Split & Deduplica) */}
                        <div className="flex-1 min-w-[300px] bg-slate-50 border-2 border-slate-200 rounded-xl px-4 py-2 flex flex-col justify-center">
                            <div className="flex items-center gap-2 mb-1">
                                <Settings2 size={14} className="text-orange-600" />
                                <span className="text-[10px] font-black text-slate-500 uppercase">Strumenti Pro</span>
                            </div>
                            <div className="flex items-center gap-4">
                                <label className="flex items-center gap-1.5 cursor-pointer text-xs font-bold text-slate-700 hover:text-orange-700" title="Divide 'Mario Rossi' in Nome e Cognome se mappi solo il Nome">
                                    <input type="checkbox" checked={splitNameEnabled} onChange={(e) => setSplitNameEnabled(e.target.checked)} className="rounded text-orange-600 focus:ring-orange-500" /> 
                                    <span className="flex items-center gap-1"><Split size={12}/> Split Nomi</span>
                                </label>
                                <label className="flex items-center gap-1.5 cursor-pointer text-xs font-bold text-slate-700 hover:text-orange-700" title="Rimuove righe con stessa Email o Telefono">
                                    <input type="checkbox" checked={deduplicateEnabled} onChange={(e) => setDeduplicateEnabled(e.target.checked)} className="rounded text-orange-600 focus:ring-orange-500" /> 
                                    <span className="flex items-center gap-1"><CopyMinus size={12}/> No Duplicati</span>
                                </label>
                            </div>
                        </div>

                         {/* PANEL: GLOBAL TAG */}
                         <div className="flex-1 min-w-[200px] bg-slate-50 border-2 border-slate-200 rounded-xl px-4 py-2 flex flex-col justify-center">
                            <div className="flex items-center gap-2 mb-1">
                                <Tag size={14} className="text-pink-600" />
                                <span className="text-[10px] font-black text-slate-500 uppercase">Tag Globale</span>
                            </div>
                            <input 
                                type="text" 
                                value={globalTag} 
                                onChange={(e) => setGlobalTag(e.target.value)} 
                                className="w-full bg-white border border-slate-300 rounded px-2 py-0.5 text-xs font-bold placeholder:font-normal outline-none focus:border-pink-400" 
                                placeholder="es. Fiera 2026"
                            />
                        </div>
                    </div>

                    <button onClick={() => { setStep(1); setFileData([]); }} className="text-slate-400 hover:text-red-600 p-3 hover:bg-red-50 rounded-xl transition-all self-center"><X size={24} /></button>
                </div>

                <div className="p-6 md:p-8 overflow-y-auto">
                    {/* ALERT DI CONFIGURAZIONE */}
                    <div className="flex flex-col md:flex-row items-center justify-between mb-8 bg-white p-5 rounded-3xl border-2 border-slate-200 shadow-md gap-4">
                        <div className="flex items-center gap-4">
                            <div className="p-3 bg-slate-100 rounded-2xl text-slate-700"><Layout size={24} /></div> 
                            <div>
                                <h3 className="font-black text-2xl text-slate-900">Mappatura Campi</h3>
                                <p className="text-xs text-slate-500 font-bold uppercase tracking-wide">
                                    {splitNameEnabled ? <span className="text-orange-600 flex items-center gap-1"><Split size={12}/> Split Nome Attivo</span> : "Modalità Standard"}
                                    {deduplicateEnabled && <span className="text-orange-600 flex items-center gap-1 ml-2"><CopyMinus size={12}/> Deduplica Attiva</span>}
                                </p>
                            </div>
                        </div>
                        <div className={`px-5 py-3 rounded-full text-sm font-bold border-2 flex items-center gap-2 shadow-sm ${isMappingValid() ? 'bg-emerald-100 text-emerald-900 border-emerald-400' : 'bg-amber-100 text-amber-900 border-amber-400'}`}>
                            {isMappingValid() ? <CheckCircle2 size={20} strokeWidth={2.5}/> : <AlertCircle size={20} strokeWidth={2.5}/>}
                            {isMappingValid() ? 'Configurazione Valida' : 'Manca Email o Telefono'}
                        </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6 mb-10">
                        {DEMO_COLUMNS.map((col) => (
                        <div key={col.id} className={`p-5 rounded-3xl border-2 transition-all shadow-md ${mapping[col.id] ? 'border-emerald-500 bg-emerald-50 shadow-emerald-200/50' : 'border-slate-300 bg-white hover:border-slate-400 hover:shadow-lg'}`}>
                            <div className="flex justify-between items-start mb-3">
                                <div className="flex gap-2">
                                  <span className={`text-xs font-black uppercase tracking-wider px-3 py-1.5 rounded-lg border flex items-center gap-2 ${getCategoryColor(col.category)}`}>
                                    {getCategoryIcon(col.category)} {col.category}
                                  </span>
                                  {/* ICONE STATUS */}
                                  {cleaningEnabled && col.cleanType !== 'none' && (
                                    <span className="text-[10px] font-bold bg-indigo-100 text-indigo-700 px-2 py-1 rounded-lg border border-indigo-200 flex items-center gap-1" title="Pulizia attiva su questo campo">
                                      <Wand2 size={10} /> Auto
                                    </span>
                                  )}
                                  {splitNameEnabled && col.id === 'nome' && (
                                     <span className="text-[10px] font-bold bg-orange-100 text-orange-700 px-2 py-1 rounded-lg border border-orange-200 flex items-center gap-1" title="Verrà diviso in Nome e Cognome">
                                     <Split size={10} /> Split
                                   </span>
                                  )}
                                </div>
                                {col.requiredGroup && <span className={`h-3 w-3 rounded-full ring-4 ring-white shadow-sm ${mapping[col.id] ? 'bg-emerald-500' : 'bg-amber-500 animate-pulse'}`} title="Campo Obbligatorio (Gruppo)" />}
                            </div>
                            
                            <p className="font-black text-xl text-slate-900 mb-1 capitalize">{col.label}</p>
                            <p className="text-sm text-slate-700 mb-4 leading-snug font-medium min-h-[1.5em]">{col.description}</p>
                            
                            <div className="relative">
                                <select 
                                  value={mapping[col.id] || ""} 
                                  onChange={(e) => setMapping(prev => ({ ...prev, [col.id]: e.target.value }))} 
                                  disabled={splitNameEnabled && col.id === 'cognome'}
                                  className={`w-full text-sm font-bold p-4 rounded-2xl border-2 outline-none appearance-none cursor-pointer transition-all shadow-sm 
                                    ${splitNameEnabled && col.id === 'cognome' ? 'bg-slate-100 text-slate-400 border-slate-200 cursor-not-allowed' : ''}
                                    ${mapping[col.id] ? 'bg-white border-emerald-600 text-emerald-800 focus:ring-4 focus:ring-emerald-200' : 'bg-slate-50 border-slate-300 text-slate-800 hover:bg-white hover:border-slate-500 focus:ring-4 focus:ring-slate-200'}`}
                                >
                                    <option value="" className="text-slate-500">
                                        {splitNameEnabled && col.id === 'cognome' ? "Auto-generato da Nome" : "-- Ignora questo campo --"}
                                    </option>
                                    {userHeaders.map(h => (
                                      <option key={h} value={h} className="text-slate-900 font-bold">
                                        {h} {headerSamples[h] ? `(es. ${headerSamples[h]})` : ''}
                                      </option>
                                    ))}
                                </select>
                                {!(splitNameEnabled && col.id === 'cognome') && <ArrowRight className={`absolute right-4 top-4 pointer-events-none transition-colors ${mapping[col.id] ? 'text-emerald-600' : 'text-slate-500'}`} size={20} strokeWidth={2.5} />}
                            </div>
                        </div>
                        ))}
                    </div>

                    {/* LIVE PREVIEW TABLE */}
                    <div className="rounded-3xl border-2 border-slate-300 overflow-hidden shadow-md bg-white mb-6">
                        <div className="bg-slate-100 px-6 py-4 border-b-2 border-slate-300 flex justify-between items-center">
                            <div className="flex items-center gap-3">
                              <span className="text-sm font-black text-slate-800 uppercase tracking-wider flex items-center gap-2"><Layout size={18} className="text-slate-600"/> Anteprima Dati (Top 30)</span>
                            </div>
                            
                            <button onClick={handleManualUpdate} className="flex items-center gap-2 bg-slate-900 text-white px-4 py-2 rounded-xl text-xs font-bold hover:bg-emerald-600 transition-colors shadow-sm">
                              <RefreshCw size={14} /> Aggiorna Anteprima
                            </button>
                        </div>
                        <div className="overflow-x-auto max-h-[400px]">
                          <table className="w-full text-left text-sm">
                            <thead className="bg-slate-200 text-slate-800 font-black uppercase border-b-2 border-slate-300 sticky top-0 z-10">
                              <tr>{DEMO_COLUMNS.map(c => <th key={c.id} className="px-6 py-4 whitespace-nowrap tracking-wide">{c.label}</th>)}</tr>
                            </thead>
                            <tbody className="divide-y-2 divide-slate-200 text-slate-900 font-bold">
                              {previewRows.map((r, i) => (
                                <tr key={i} className="hover:bg-slate-100 transition-colors">
                                  {DEMO_COLUMNS.map(c => (
                                    <td key={c.id} className={`px-6 py-4 whitespace-nowrap ${r[c.label] === '-' || r[c.label] === '' ? 'text-slate-400 font-medium italic' : ''}`}>
                                      {r[c.label] || "-"}
                                    </td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                    </div>
                </div>
            </div>
            )}
        </div>
      </div>

      {step === 2 && (
        <div className="fixed bottom-0 left-0 w-full bg-white/95 backdrop-blur-md border-t-4 border-slate-200 p-5 z-40 shadow-[0_-10px_30px_rgba(0,0,0,0.15)]">
            <div className="max-w-6xl mx-auto flex flex-col sm:flex-row justify-between items-center gap-5">
                <div className="hidden md:flex items-center gap-3 w-full max-w-md">
                     <p className="text-xs font-black text-slate-500 uppercase tracking-wider mb-1">Salva come</p>
                     <input type="text" value={customFileName} onChange={(e) => setCustomFileName(e.target.value)} className="w-full bg-slate-100 border-2 border-slate-300 rounded-lg px-3 py-2 text-sm font-bold text-slate-900 outline-none focus:border-emerald-500" placeholder="Nome file..." />
                </div>

                <button onClick={exportFile} disabled={!isMappingValid()} className={`w-full md:w-auto px-10 py-4 rounded-2xl font-black text-white shadow-xl transition-all transform active:scale-95 flex items-center justify-center gap-3 text-xl ${isMappingValid() ? 'bg-gradient-to-br from-slate-900 to-slate-800 hover:from-emerald-600 hover:to-emerald-800 hover:shadow-emerald-500/40 border-4 border-transparent' : 'bg-slate-400 text-slate-600 cursor-not-allowed border-4 border-slate-300'}`}>
                    <Download size={24} strokeWidth={3} /> Scarica CSV Pienissimo
                </button>
            </div>
        </div>
      )}

      <footer className="mt-auto py-8 text-center text-slate-600 font-bold bg-slate-100 border-t-2 border-slate-200">
          <div className="mb-2"><span className="opacity-70 text-xs uppercase tracking-[0.2em] font-black text-slate-500">Powered by</span></div>
          <p className="text-lg font-black text-slate-900">Nicola Pellicioni</p>
          <div className="flex items-center justify-center gap-2 mt-1 text-sm"><span className="text-slate-700">Responsabile Assistenza Tecnica</span><span className="h-2 w-2 bg-emerald-600 rounded-full box-shadow-sm"></span><span className="font-black text-emerald-800 uppercase">Pienissimo Software</span></div>
          <p className="text-xs mt-4 font-mono bg-slate-200 inline-block px-4 py-1.5 rounded-full border border-slate-300 text-slate-700">v1.0.0 • Internal Tool</p>
      </footer>
    </div>
  );
}