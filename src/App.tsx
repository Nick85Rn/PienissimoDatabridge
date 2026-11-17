import React, { useState, type ChangeEvent } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import {
  Upload, Download, CheckCircle2, AlertCircle, FileSpreadsheet, ArrowRight,
  Loader2, X, ShieldCheck, MapPin, User, Tag, RefreshCw,
  Wand2, Split, CopyMinus, Stamp, ScrollText, BadgeCheck
} from 'lucide-react';
import { IMPORT_COLUMNS as DEMO_COLUMNS, autoMatch, processRows, toXlsxCells } from './importFormat';
import type { Mapping, Row, OutRow, ImportStats } from './importFormat';

export default function App() {
  const [step, setStep] = useState<1 | 2>(1);
  const [fileData, setFileData] = useState<Row[]>([]);
  const [userHeaders, setUserHeaders] = useState<string[]>([]);
  const [headerSamples, setHeaderSamples] = useState<Record<string, string>>({});
  const [mapping, setMapping] = useState<Mapping>({});

  // --- STATI AVANZATI ---
  const [cleaningEnabled, setCleaningEnabled] = useState<boolean>(true);
  const [splitNameEnabled, setSplitNameEnabled] = useState<boolean>(false);
  const [deduplicateEnabled, setDeduplicateEnabled] = useState<boolean>(false);
  const [globalTag, setGlobalTag] = useState<string>("");
  const [defaultPrefix, setDefaultPrefix] = useState<string>("+39");

  const [previewRows, setPreviewRows] = useState<OutRow[]>([]);
  const [originalFileName, setOriginalFileName] = useState<string>("");
  const [customFileName, setCustomFileName] = useState<string>("");
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showSuccessToast, setShowSuccessToast] = useState(false);
  const [encodingWarning, setEncodingWarning] = useState(false);
  const [importStats, setImportStats] = useState<ImportStats>({ totale: 0, duplicatiRimossi: 0, senzaContatto: 0, dateNonValide: 0, consensiNonRiconosciuti: 0, emailNonValide: 0 });

  // --- ELABORAZIONE (logica in importFormat.ts) ---
  const processAllData = (dataSource: Row[], currentMapping: Mapping) => {
    const { rows, stats } = processRows(dataSource, currentMapping, {
      cleaningEnabled, splitNameEnabled, deduplicateEnabled, globalTag, defaultPrefix,
    });
    setImportStats(stats);
    return rows;
  };

  // --- RILEVAMENTO ENCODING SOSPETTO ---
  // I CSV esportati da Excel italiano sono spesso in Windows-1252/Latin-1.
  // Se vengono letti come UTF-8, lettere accentate ed Euro diventano sequenze
  // "mojibake" tipiche (es. città -> citt , perché -> perchÃ©). Qui controlliamo
  // un campione di valori in cerca di questi pattern, solo per avvisare
  // l'utente: non blocchiamo né correggiamo automaticamente, perché un fix
  // automatico sbagliato rischierebbe di rovinare dati già corretti.
  const looksLikeMojibake = (rows: Row[], headers: string[]) => {
    const mojibakePattern = /Ã[\u0080-\u00BF]|Â[\u0080-\u00BF]|â€™|â€œ|â€/;
    let sampleCount = 0;
    for (const row of rows) {
      for (const h of headers) {
        const val = row[h];
        if (typeof val === 'string' && val) {
          if (mojibakePattern.test(val)) return true;
          sampleCount++;
        }
      }
      if (sampleCount > 200) break;
    }
    return false;
  };

  // --- LOGICA DI IMPORT ---
  const processFile = (file: File) => {
    setIsProcessing(true);
    setOriginalFileName(file.name);
    setCustomFileName(file.name.split('.')[0]);
    const extension = file.name.split('.').pop()?.toLowerCase();

    const parseCallback = (data: Row[], headers: string[]) => {
      setTimeout(() => {
        setUserHeaders(headers);
        setFileData(data);
        setEncodingWarning(looksLikeMojibake(data, headers));

        // Esempio per ogni colonna: primo valore non vuoto di quella colonna
        const samples: Record<string, string> = {};
        headers.forEach(h => {
          const row = data.find(r => r[h] !== undefined && String(r[h]).trim() !== '');
          const val = row ? String(row[h]) : '';
          samples[h] = val ? val.substring(0, 15) + (val.length > 15 ? '...' : '') : '';
        });
        setHeaderSamples(samples);

        const initialMapping = autoMatch(headers);
        setMapping(initialMapping);

        // Genera anteprima iniziale
        const initialProcessed = processAllData(data, initialMapping);
        setPreviewRows(initialProcessed.slice(0, 30));

        setStep(2);
        setIsProcessing(false);
      }, 800);
    };

    if (extension === 'csv') {
      Papa.parse<Row>(file, {
        header: true,
        skipEmptyLines: 'greedy',
        complete: (results) => parseCallback(results.data, (results.meta.fields || []).filter(f => f.trim() !== ''))
      });
    } else if (['xlsx', 'xls'].includes(extension || '')) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const wb = XLSX.read(evt.target?.result as ArrayBuffer, { type: 'array' });
          const ws = wb.Sheets[wb.SheetNames[0]];
          // Intestazioni lette dalla prima riga, così compaiono anche le colonne vuote nei primi record
          const headerRow = (XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, blankrows: false })[0] || [])
            .map(h => String(h ?? '').trim()).filter(h => h !== '');
          // raw: true -> date come numero seriale Excel, convertite poi in AAAA-MM-GG
          const data = XLSX.utils.sheet_to_json<Row>(ws, { defval: '', raw: true, blankrows: false });
          parseCallback(data, headerRow);
        } catch {
          setIsProcessing(false);
          alert("Impossibile leggere il file Excel (protetto da password o danneggiato?).");
        }
      };
      reader.readAsArrayBuffer(file);
    } else { setIsProcessing(false); alert("Formato non supportato."); }
  };

  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>) => { const file = e.target.files?.[0]; if (file) processFile(file); };
  const onDrag = (e: React.DragEvent, dragging: boolean) => { e.preventDefault(); setIsDragging(dragging); };
  const onDrop = (e: React.DragEvent) => { e.preventDefault(); setIsDragging(false); const file = e.dataTransfer.files?.[0]; if (file) processFile(file); };

  const isMappingValid = () => !!(mapping['email'] || mapping['telefono']);

  const handleManualUpdate = () => {
    const fullProcessed = processAllData(fileData, mapping);
    setPreviewRows(fullProcessed.slice(0, 30));
  };

  // --- ESPORTAZIONE ---
  const exportFile = (format: 'xlsx' | 'csv') => {
    const fullProcessed = processAllData(fileData, mapping);

    const now = new Date();
    const p2 = (n: number) => n.toString().padStart(2, '0');
    const ts = `${now.getFullYear()}${p2(now.getMonth()+1)}${p2(now.getDate())}_${p2(now.getHours())}${p2(now.getMinutes())}`;
    const baseName = `${customFileName || 'importazione'}_${ts}`;

    if (format === 'xlsx') {
      // Stesso formato del modello "esempio_importazioni.xlsx": foglio "Foglio1", colonne nello stesso ordine
      const ws = XLSX.utils.aoa_to_sheet(toXlsxCells(fullProcessed));
      ws['!cols'] = DEMO_COLUMNS.map(c => ({ wch: Math.max(12, c.label.length + 2, c.id === 'email' ? 30 : 0) }));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Foglio1');
      XLSX.writeFile(wb, `${baseName}.xlsx`);
    } else {
      const link = document.createElement("a");
      // BOM iniziale: Excel apre correttamente gli accenti
      const csv = Papa.unparse({ fields: DEMO_COLUMNS.map(c => c.label), data: fullProcessed.map(r => DEMO_COLUMNS.map(c => r[c.label] ?? '')) });
      link.href = URL.createObjectURL(new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' }));
      link.download = `${baseName}.csv`;
      link.click();
      setTimeout(() => URL.revokeObjectURL(link.href), 1000);
    }
    setShowSuccessToast(true); setTimeout(() => setShowSuccessToast(false), 4000);
  };

  const getCategoryIcon = (cat: string) => {
    switch(cat) {
      case 'Anagrafica': return <User size={13} strokeWidth={2.25} />;
      case 'Contatti': return <ShieldCheck size={13} strokeWidth={2.25} />;
      case 'Indirizzo': return <MapPin size={13} strokeWidth={2.25} />;
      case 'Dettagli': return <Tag size={13} strokeWidth={2.25} />;
      case 'Consensi': return <BadgeCheck size={13} strokeWidth={2.25} />;
      default: return null;
    }
  };

  // Conteggio campi mappati per il "codice lotto" e per la tassonomia
  const mappedCount = DEMO_COLUMNS.filter(c => mapping[c.id]).length;
  const lotCode = originalFileName
    ? originalFileName.replace(/\.[^/.]+$/, '').slice(0, 10).toUpperCase().replace(/[^A-Z0-9]/g, '-')
    : 'IN-ATTESA';

  return (
    <div className="min-h-screen bg-paper bg-paper-grain bg-[size:18px_18px] text-ink font-sans pb-40 flex flex-col justify-between">

      {/* TOAST CONFORME */}
      {showSuccessToast && (
        <div className="fixed top-6 right-6 z-[100] animate-in slide-in-from-top-4 fade-in duration-300">
          <div className="bg-paper border-2 border-ink/80 shadow-[3px_3px_0_0_rgba(28,43,36,0.9)] rounded-sm p-4 flex items-start gap-3 max-w-sm">
            <div className="border-2 border-stamp text-stamp p-1.5 rounded-full -rotate-6"><Stamp size={18} /></div>
            <div>
              <h4 className="font-mono font-bold text-sm uppercase tracking-wide text-stamp-dark">Lotto Esportato</h4>
              <p className="text-sm text-ink-soft mt-1">File pronto per l'importazione in Pienissimo.</p>
            </div>
            <button onClick={() => setShowSuccessToast(false)} className="text-ink-soft/50 hover:text-ink ml-2"><X size={18} /></button>
          </div>
        </div>
      )}

      <div className="max-w-[95%] xl:max-w-6xl mx-auto px-4 w-full py-8">

        {/* HEADER — fascia "etichetta di spedizione" */}
        <header className="relative mb-8 bg-ink rounded-sm shadow-[4px_4px_0_0_rgba(28,43,36,0.18)]">
          <div className="flex flex-col md:flex-row items-stretch justify-between">
            <div className="flex items-center gap-4 px-7 py-6 border-b md:border-b-0 md:border-r border-paper/15">
              <div className="border-2 border-stamp-bright/70 rounded-full p-2.5 text-stamp-bright">
                <ShieldCheck size={26} strokeWidth={2} />
              </div>
              <div>
                <h1 className="text-xl font-extrabold text-paper tracking-tight leading-none">Pienissimo</h1>
                <span className="text-[11px] font-mono font-semibold text-stamp-bright tracking-[0.18em] uppercase">DataBridge — Banco Controllo</span>
              </div>
            </div>
            <div className="flex items-center gap-5 px-7 py-4 md:py-0">
              <div className="font-mono text-[11px] text-paper/60 leading-tight">
                <p className="uppercase tracking-[0.15em]">Lotto</p>
                <p className="text-paper font-bold text-sm tracking-wide">{lotCode}</p>
              </div>
              <div className="h-8 w-px bg-paper/15" />
              <div className="font-mono text-[11px] text-paper/60 leading-tight">
                <p className="uppercase tracking-[0.15em]">Stato</p>
                <p className={`font-bold text-sm tracking-wide ${step === 2 ? 'text-stamp-bright' : 'text-paper/50'}`}>{step === 2 ? 'In Ispezione' : 'In Attesa'}</p>
              </div>
            </div>
          </div>
          {/* bordo perforato in basso, come una ricevuta da strappare */}
          <div className="h-2 w-full overflow-hidden flex items-center px-3" aria-hidden>
            <div className="w-full border-b-2 border-dotted border-paper/15" />
          </div>
        </header>

        {isProcessing && (
          <div className="fixed inset-0 bg-ink/85 backdrop-blur-sm z-50 flex flex-col items-center justify-center text-paper">
            <Loader2 className="animate-spin text-stamp mb-4" size={56} strokeWidth={2} />
            <p className="text-xl font-mono font-bold uppercase tracking-wide">Verifica del lotto in corso</p>
            <p className="text-paper/60 mt-2 font-medium text-sm">Lettura struttura file, rilevamento colonne…</p>
          </div>
        )}

        <div className="bg-paper rounded-sm shadow-[4px_4px_0_0_rgba(28,43,36,0.10)] border border-ink/10 overflow-hidden min-h-[600px]">

          {/* STEP 1 — RICEZIONE FILE */}
          {step === 1 && (
            <div className="h-full flex flex-col items-center justify-center p-8 md:p-16 animate-in fade-in zoom-in-95 duration-500">
              <div className="text-center mb-10 max-w-lg">
                <p className="font-mono text-xs uppercase tracking-[0.2em] text-stamp font-bold mb-3">Modulo di Ricezione</p>
                <h3 className="text-3xl md:text-4xl font-extrabold text-ink mb-4 tracking-tight">Carica il database clienti</h3>
                <p className="text-ink-soft text-base leading-relaxed">
                  Excel o CSV, qualsiasi struttura. Verrà letto, controllato e convertito nel formato standard pronto per Pienissimo.
                </p>
              </div>
              <label onDragOver={(e) => onDrag(e, true)} onDragLeave={(e) => onDrag(e, false)} onDrop={onDrop}
                className={`group relative flex flex-col items-center justify-center w-full max-w-2xl h-64 rounded-sm cursor-pointer transition-all duration-200 ${isDragging ? 'bg-stamp-light border-2 border-stamp scale-[1.01]' : 'bg-paper-dark border-2 border-dashed border-ink/25 hover:border-stamp/60 hover:bg-stamp-light/40'}`}>
                <div className={`rounded-full p-4 mb-4 border-2 transition-colors ${isDragging ? 'border-stamp text-stamp bg-paper' : 'border-ink/20 text-ink-soft group-hover:border-stamp/50 group-hover:text-stamp'}`}>
                  <Upload size={32} strokeWidth={2} />
                </div>
                <h4 className="font-bold text-lg text-ink">{isDragging ? 'Rilascia qui il file' : 'Clicca o trascina il file'}</h4>
                <p className="text-xs mt-3 font-mono uppercase tracking-wider text-ink-soft/70 border border-ink/15 rounded-full px-4 py-1.5">CSV · XLS · XLSX</p>
                <input type="file" className="hidden" accept=".csv, .xls, .xlsx" onChange={handleFileUpload} />
              </label>
            </div>
          )}

          {/* STEP 2 — ISPEZIONE E MAPPATURA */}
          {step === 2 && (
            <div className="animate-in slide-in-from-right-8 duration-500 flex flex-col h-full">

              {/* BARRA STRUMENTI */}
              <div className="bg-paper-dark border-b border-ink/10 p-4 md:px-6 flex flex-col xl:flex-row items-start xl:items-center justify-between gap-4 sticky top-0 z-10">
                <div className="flex flex-wrap gap-3 w-full">

                  {/* FILE INFO */}
                  <div className="flex items-center gap-3 bg-paper border border-ink/15 rounded-sm px-3.5 py-2">
                    <FileSpreadsheet size={18} className="text-stamp" strokeWidth={2} />
                    <div className="max-w-[140px]">
                      <p className="text-[9px] font-mono uppercase tracking-wider text-ink-soft/60">File</p>
                      <p className="font-semibold text-ink text-xs truncate">{originalFileName}</p>
                    </div>
                  </div>

                  {/* PULIZIA */}
                  <div className="flex-1 min-w-[260px] bg-paper border border-ink/15 rounded-sm px-4 py-2 flex flex-col justify-center">
                    <div className="flex items-center gap-1.5 mb-1">
                      <Wand2 size={12} className="text-stamp" strokeWidth={2} />
                      <span className="text-[9px] font-mono font-bold text-ink-soft/60 uppercase tracking-wider">Pulizia Dati</span>
                    </div>
                    <div className="flex items-center gap-3">
                      <label className="flex items-center gap-1.5 cursor-pointer text-xs font-semibold text-ink">
                        <input type="checkbox" checked={cleaningEnabled} onChange={(e) => setCleaningEnabled(e.target.checked)} className="rounded-sm text-stamp focus:ring-stamp/40 accent-stamp" /> Auto-Fix
                      </label>
                      {cleaningEnabled && (
                        <input type="text" value={defaultPrefix} onChange={(e) => setDefaultPrefix(e.target.value)} className="w-16 px-1.5 py-0.5 text-xs font-mono font-semibold border border-ink/20 rounded-sm text-center bg-paper" placeholder="+39" title="Prefisso Default" />
                      )}
                    </div>
                  </div>

                  {/* STRUMENTI PRO */}
                  <div className="flex-1 min-w-[280px] bg-paper border border-ink/15 rounded-sm px-4 py-2 flex flex-col justify-center">
                    <div className="flex items-center gap-1.5 mb-1">
                      <Stamp size={12} className="text-rust" strokeWidth={2} />
                      <span className="text-[9px] font-mono font-bold text-ink-soft/60 uppercase tracking-wider">Strumenti Pro</span>
                    </div>
                    <div className="flex items-center gap-4">
                      <label className="flex items-center gap-1.5 cursor-pointer text-xs font-semibold text-ink" title="Divide 'Mario Rossi' in Nome e Cognome se mappi solo il Nome">
                        <input type="checkbox" checked={splitNameEnabled} onChange={(e) => setSplitNameEnabled(e.target.checked)} className="rounded-sm accent-rust focus:ring-rust/40" />
                        <span className="flex items-center gap-1"><Split size={12}/> Split Nomi</span>
                      </label>
                      <label className="flex items-center gap-1.5 cursor-pointer text-xs font-semibold text-ink" title="Rimuove righe con stessa Email o Telefono">
                        <input type="checkbox" checked={deduplicateEnabled} onChange={(e) => setDeduplicateEnabled(e.target.checked)} className="rounded-sm accent-rust focus:ring-rust/40" />
                        <span className="flex items-center gap-1"><CopyMinus size={12}/> No Duplicati</span>
                      </label>
                    </div>
                  </div>

                  {/* TAG GLOBALE */}
                  <div className="flex-1 min-w-[190px] bg-paper border border-ink/15 rounded-sm px-4 py-2 flex flex-col justify-center">
                    <div className="flex items-center gap-1.5 mb-1">
                      <Tag size={12} className="text-ink-soft" strokeWidth={2} />
                      <span className="text-[9px] font-mono font-bold text-ink-soft/60 uppercase tracking-wider">Tag Globale</span>
                    </div>
                    <input
                      type="text"
                      value={globalTag}
                      onChange={(e) => setGlobalTag(e.target.value)}
                      className="w-full bg-transparent border-b border-ink/20 px-0.5 py-0.5 text-xs font-semibold placeholder:font-normal placeholder:text-ink-soft/40 outline-none focus:border-stamp"
                      placeholder="es. Fiera 2026"
                    />
                  </div>
                </div>

                <button onClick={() => { setStep(1); setFileData([]); setEncodingWarning(false); }} className="text-ink-soft/50 hover:text-rust p-2.5 hover:bg-rust-light rounded-sm transition-all self-center" title="Annulla e ricomincia">
                  <X size={20} />
                </button>
              </div>

              <div className="p-6 md:p-8 overflow-y-auto">

                {/* REPORT DI ISPEZIONE */}
                {(deduplicateEnabled && (importStats.duplicatiRimossi > 0 || importStats.senzaContatto > 0)) && (
                  <div className="flex flex-wrap items-stretch gap-0 mb-5 bg-paper border border-ink/15 rounded-sm overflow-hidden">
                    <div className="flex items-center gap-2 px-4 py-3 bg-ink/5 border-r border-ink/10">
                      <ScrollText size={16} className="text-ink-soft" />
                      <span className="text-[10px] font-mono font-bold uppercase tracking-wider text-ink-soft">Esito Deduplica</span>
                    </div>
                    <div className="flex flex-wrap gap-6 px-5 py-2.5 items-center">
                      <div>
                        <p className="text-[9px] font-mono uppercase tracking-wider text-ink-soft/60">Righe Totali</p>
                        <p className="font-mono font-bold text-base text-ink">{importStats.totale}</p>
                      </div>
                      <div>
                        <p className="text-[9px] font-mono uppercase tracking-wider text-ink-soft/60">Duplicati Rimossi</p>
                        <p className="font-mono font-bold text-base text-rust">{importStats.duplicatiRimossi}</p>
                      </div>
                      <div>
                        <p className="text-[9px] font-mono uppercase tracking-wider text-ink-soft/60">Senza Contatto</p>
                        <p className="font-mono font-bold text-base text-stamp-dark">{importStats.senzaContatto}</p>
                        <p className="text-[10px] text-ink-soft/60">mantenute, non deduplicabili</p>
                      </div>
                    </div>
                  </div>
                )}

                {/* CONTROLLI FORMATO IMPORTAZIONE */}
                {(importStats.dateNonValide > 0 || importStats.consensiNonRiconosciuti > 0 || importStats.emailNonValide > 0) && (
                  <div className="flex items-start gap-3 mb-5 bg-rust-light p-4 rounded-sm border border-rust/30">
                    <AlertCircle size={20} className="text-rust mt-0.5 flex-shrink-0" />
                    <div>
                      <p className="font-bold text-sm text-rust-dark">Valori da controllare prima dell'importazione</p>
                      <ul className="text-sm text-ink-soft mt-1 leading-snug list-disc pl-4">
                        {importStats.dateNonValide > 0 && <li><b>{importStats.dateNonValide}</b> date di nascita non riconosciute: lasciate come nel file originale.</li>}
                        {importStats.consensiNonRiconosciuti > 0 && <li><b>{importStats.consensiNonRiconosciuti}</b> consensi con valore non riconosciuto: impostati a 0 (nessun consenso).</li>}
                        {importStats.emailNonValide > 0 && <li><b>{importStats.emailNonValide}</b> email con formato non valido.</li>}
                      </ul>
                    </div>
                  </div>
                )}

                {/* AVVISO ENCODING */}
                {encodingWarning && (
                  <div className="flex items-start gap-3 mb-5 bg-rust-light p-4 rounded-sm border border-rust/30">
                    <AlertCircle size={20} className="text-rust mt-0.5 flex-shrink-0" />
                    <div>
                      <p className="font-bold text-sm text-rust-dark">Possibile problema di codifica caratteri</p>
                      <p className="text-sm text-ink-soft mt-1 leading-snug">Alcuni valori contengono caratteri anomali (lettere accentate trasformate in simboli strani): il file è probabilmente salvato in una codifica diversa da UTF-8. Controlla l'anteprima qui sotto — se città o vie sono illeggibili, riesporta il CSV da Excel scegliendo "CSV UTF-8".</p>
                    </div>
                  </div>
                )}

                {/* TIMBRO DI CONFORMITÀ + INTESTAZIONE MAPPATURA */}
                <div className="flex flex-col md:flex-row items-center justify-between mb-6 gap-4 pb-5 border-b-2 border-dotted border-ink/15">
                  <div>
                    <h3 className="font-extrabold text-xl text-ink tracking-tight">Mappatura Campi</h3>
                    <p className="text-xs text-ink-soft font-mono uppercase tracking-wide mt-1 flex items-center gap-3 flex-wrap">
                      <span>{mappedCount} / {DEMO_COLUMNS.length} campi mappati</span>
                      {splitNameEnabled && <span className="text-rust flex items-center gap-1"><Split size={11}/> Split Attivo</span>}
                      {deduplicateEnabled && <span className="text-rust flex items-center gap-1"><CopyMinus size={11}/> Deduplica Attiva</span>}
                    </p>
                  </div>

                  {/* Timbro: ruotato, doppio bordo, effetto "impresso" */}
                  <div
                    className={`relative px-5 py-2.5 rounded-sm border-[3px] font-mono font-bold text-sm uppercase tracking-wider flex items-center gap-2 -rotate-2 transition-all
                      ${isMappingValid()
                        ? 'border-stamp text-stamp-dark bg-stamp-light/60'
                        : 'border-rust text-rust-dark bg-rust-light/50 animate-pulse'}`}
                    style={{ boxShadow: isMappingValid() ? '2px 2px 0 0 rgba(47,107,79,0.35)' : '2px 2px 0 0 rgba(181,72,42,0.3)' }}
                  >
                    {isMappingValid() ? <CheckCircle2 size={18} strokeWidth={2.5}/> : <AlertCircle size={18} strokeWidth={2.5}/>}
                    {isMappingValid() ? 'Conforme' : 'Manca Contatto'}
                  </div>
                </div>

                {/* GRID CAMPI */}
                <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4 mb-10">
                  {DEMO_COLUMNS.map((col) => (
                    <div key={col.id} className={`relative p-4 rounded-sm border transition-all ${mapping[col.id] ? 'border-stamp/50 bg-stamp-light/30' : 'border-ink/15 bg-paper hover:border-ink/30'}`}>
                      <div className="flex justify-between items-start mb-2.5">
                        <div className="flex gap-1.5 flex-wrap">
                          <span className="text-[10px] font-mono font-bold uppercase tracking-wider px-2 py-1 rounded-sm border border-ink/15 text-ink-soft flex items-center gap-1.5 bg-paper">
                            {getCategoryIcon(col.category)} {col.category}
                          </span>
                          {cleaningEnabled && col.cleanType !== 'none' && (
                            <span className="text-[9px] font-mono font-bold text-stamp-dark px-1.5 py-1 rounded-sm border border-stamp/30 bg-stamp-light/50 flex items-center gap-1" title="Pulizia attiva su questo campo">
                              <Wand2 size={9} /> AUTO
                            </span>
                          )}
                          {splitNameEnabled && col.id === 'nome' && (
                            <span className="text-[9px] font-mono font-bold text-rust-dark px-1.5 py-1 rounded-sm border border-rust/30 bg-rust-light/50 flex items-center gap-1" title="Verrà diviso in Nome e Cognome">
                              <Split size={9} /> SPLIT
                            </span>
                          )}
                        </div>
                        {col.requiredGroup && (
                          <span className={`h-2.5 w-2.5 rounded-full flex-shrink-0 mt-1 ${mapping[col.id] ? 'bg-stamp' : 'bg-rust animate-pulse'}`} title="Campo del gruppo obbligatorio (email o telefono)" />
                        )}
                      </div>

                      <p className="font-bold text-base text-ink mb-1 capitalize">{col.label}</p>
                      <p className="text-xs text-ink-soft mb-3 leading-snug min-h-[2em]">{col.description}</p>

                      <div className="relative">
                        <select
                          value={mapping[col.id] || ""}
                          onChange={(e) => setMapping(prev => ({ ...prev, [col.id]: e.target.value }))}
                          disabled={splitNameEnabled && col.id === 'cognome'}
                          className={`w-full text-sm font-semibold p-3 pr-9 rounded-sm border outline-none appearance-none cursor-pointer transition-all
                            ${splitNameEnabled && col.id === 'cognome' ? 'bg-ink/5 text-ink-soft/50 border-ink/10 cursor-not-allowed' : ''}
                            ${mapping[col.id] && !(splitNameEnabled && col.id === 'cognome') ? 'bg-paper border-stamp text-stamp-dark focus:ring-2 focus:ring-stamp/25' : ''}
                            ${!mapping[col.id] && !(splitNameEnabled && col.id === 'cognome') ? 'bg-paper border-ink/20 text-ink focus:ring-2 focus:ring-ink/15' : ''}`}
                        >
                          <option value="" className="text-ink-soft">
                            {splitNameEnabled && col.id === 'cognome' ? "Auto-generato da Nome" : "— Ignora questo campo —"}
                          </option>
                          {userHeaders.map(h => (
                            <option key={h} value={h} className="text-ink font-semibold">
                              {h} {headerSamples[h] ? `(es. ${headerSamples[h]})` : ''}
                            </option>
                          ))}
                        </select>
                        {!(splitNameEnabled && col.id === 'cognome') && <ArrowRight className={`absolute right-3 top-3.5 pointer-events-none transition-colors ${mapping[col.id] ? 'text-stamp' : 'text-ink-soft/40'}`} size={16} strokeWidth={2.5} />}
                      </div>
                    </div>
                  ))}
                </div>

                {/* ANTEPRIMA — foglio a righe da registro */}
                <div className="rounded-sm border border-ink/15 overflow-hidden bg-paper mb-6">
                  <div className="bg-ink px-5 py-3 flex justify-between items-center">
                    <span className="text-xs font-mono font-bold text-paper uppercase tracking-wider flex items-center gap-2">
                      <ScrollText size={15} className="text-stamp-bright"/> Registro Anteprima — primi 30 record
                    </span>
                    <button onClick={handleManualUpdate} className="flex items-center gap-1.5 bg-stamp text-paper px-3.5 py-1.5 rounded-sm text-xs font-mono font-bold uppercase tracking-wide hover:bg-stamp-dark transition-colors">
                      <RefreshCw size={12} /> Aggiorna
                    </button>
                  </div>
                  <div className="overflow-x-auto max-h-[400px]">
                    <table className="w-full text-left text-sm">
                      <thead className="bg-paper-dark text-ink-soft font-mono font-bold uppercase border-b border-ink/15 sticky top-0 z-[5]">
                        <tr>{DEMO_COLUMNS.map(c => <th key={c.id} className="px-5 py-3 whitespace-nowrap tracking-wide text-[11px]">{c.label}</th>)}</tr>
                      </thead>
                      <tbody className="divide-y divide-ink/8 text-ink">
                        {previewRows.map((r, i) => (
                          <tr key={i} className="hover:bg-stamp-light/30 transition-colors">
                            {DEMO_COLUMNS.map(c => (
                              <td key={c.id} className={`px-5 py-3 whitespace-nowrap font-medium ${r[c.label] === '' || r[c.label] === undefined ? 'text-ink-soft/40 italic font-normal' : ''}`}>
                                {r[c.label] === '' || r[c.label] === undefined ? "—" : String(r[c.label])}
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

      {/* BARRA AZIONE FISSA — tagliando da staccare */}
      {step === 2 && (
        <div className="fixed bottom-0 left-0 w-full bg-paper border-t-2 border-dotted border-ink/20 p-4 z-40 shadow-[0_-4px_0_0_rgba(28,43,36,0.04)]">
          <div className="max-w-5xl mx-auto flex flex-col sm:flex-row justify-between items-center gap-4">
            <div className="hidden md:flex items-center gap-3 w-full max-w-sm">
              <p className="text-[10px] font-mono font-bold text-ink-soft uppercase tracking-wider whitespace-nowrap">Nome File</p>
              <input type="text" value={customFileName} onChange={(e) => setCustomFileName(e.target.value)} className="w-full bg-paper-dark border border-ink/20 rounded-sm px-3 py-2 text-sm font-semibold text-ink outline-none focus:border-stamp" placeholder="Nome file..." />
            </div>

            <div className="flex self-stretch sm:self-auto gap-3">
              <button onClick={() => exportFile('csv')} disabled={!isMappingValid()} className={`px-5 py-3.5 rounded-sm font-bold text-sm border-2 transition-all flex items-center justify-center gap-2
                ${isMappingValid() ? 'border-ink/30 text-ink hover:border-stamp hover:text-stamp-dark' : 'border-ink/15 text-ink/30 cursor-not-allowed'}`}>
                CSV
              </button>
              <button onClick={() => exportFile('xlsx')} disabled={!isMappingValid()} className={`flex-1 sm:flex-none whitespace-nowrap px-6 sm:px-8 py-3.5 rounded-sm font-bold text-paper transition-all transform active:scale-[0.98] flex items-center justify-center gap-3 text-base
                ${isMappingValid() ? 'bg-stamp hover:bg-stamp-dark shadow-[3px_3px_0_0_rgba(28,43,36,0.25)]' : 'bg-ink/25 cursor-not-allowed shadow-none'}`}>
                <Download size={20} strokeWidth={2.25} /> Esporta Excel Pienissimo
              </button>
            </div>
          </div>
        </div>
      )}

      <footer className="mt-auto py-7 text-center bg-paper-dark border-t border-ink/10">
        <p className="text-[10px] uppercase tracking-[0.2em] font-mono font-bold text-ink-soft/50 mb-1.5">Pienissimo Software · Strumento Interno</p>
        <p className="text-sm font-bold text-ink">Nicola Pellicioni</p>
        <p className="text-xs text-ink-soft mt-0.5">Responsabile Assistenza Tecnica</p>
        <p className="text-[10px] mt-3 font-mono inline-block border border-ink/15 px-3 py-1 rounded-sm text-ink-soft/70">v1.2.0</p>
      </footer>
    </div>
  );
}
