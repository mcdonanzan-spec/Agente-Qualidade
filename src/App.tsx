import { useState, useRef } from 'react';
import { GoogleGenAI } from "@google/genai";
import ReactMarkdown from 'react-markdown';
import * as pdfjsLib from 'pdfjs-dist';
import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import html2pdf from 'html2pdf.js';
import { 
  ShieldCheck, 
  FileText, 
  ClipboardCheck, 
  AlertCircle, 
  Loader2, 
  Upload,
  CheckCircle2,
  HardHat,
  Search,
  Building2,
  FileUp,
  X,
  Download,
  Printer,
  MessageSquare,
  Send
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

// Configure PDF.js worker
// Senior Dev: Using standard URL constructor allows Vite to resolve the worker path while satisfying TypeScript
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  'pdfjs-dist/build/pdf.worker.mjs',
  import.meta.url
).toString();

// Initialize Gemini API
const genAI = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

type Standard = 'ISO 9001' | 'SiAC (PBQP-H)' | 'Both';

interface ChatMessage {
  role: 'user' | 'assistant';
  content: string;
}

export default function App() {
  const [inputText, setInputText] = useState('');
  const [selectedStandard, setSelectedStandard] = useState<Standard>('Both');
  const [analysis, setAnalysis] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [isParsing, setIsParsing] = useState(false);
  const [parsingProgress, setParsingProgress] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [chatHistory, setChatHistory] = useState<ChatMessage[]>([]);
  const [userQuestion, setUserQuestion] = useState('');
  const [isChatting, setIsChatting] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'history' | 'settings'>('dashboard');
  const [auditType, setAuditType] = useState<'diagnostic' | 'internal' | 'pre-cert'>('diagnostic');
  const [confidence, setConfidence] = useState<number>(0);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const chatEndRef = useRef<HTMLDivElement>(null);
  const reportRef = useRef<HTMLDivElement>(null);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    
    setFileName(file.name);
    setIsParsing(true);
    setParsingProgress('Iniciando...');
    setError(null);

    try {
      let text = '';
      const extension = file.name.split('.').pop()?.toLowerCase();

      if (extension === 'pdf') {
        setParsingProgress('Abrindo PDF...');
        const arrayBuffer = await file.arrayBuffer();
        const loadingTask = pdfjsLib.getDocument({ 
          data: arrayBuffer,
          useSystemFonts: true,
          disableFontFace: false
        });
        const pdf = await loadingTask.promise;
        let fullText = '';
        const totalPages = pdf.numPages;
        
        for (let i = 1; i <= totalPages; i++) {
          setParsingProgress(`Lendo página ${i} de ${totalPages}...`);
          try {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            const pageText = content.items.map((item: any) => 
              'str' in item ? item.str : ''
            ).join(' ');
            fullText += pageText + '\n';
          } catch (pageErr) {
            console.warn(`Erro na página ${i}:`, pageErr);
            fullText += `[Erro na extração da página ${i}]\n`;
          }
        }
        text = fullText;
      } else if (extension === 'docx' || extension === 'doc') {
        setParsingProgress('Convertendo Word...');
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.extractRawText({ arrayBuffer });
        text = result.value;
      } else if (extension === 'xlsx' || extension === 'xls' || extension === 'csv') {
        setParsingProgress('Lendo Planilha...');
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer);
        let fullText = '';
        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          fullText += `--- Planilha: ${sheetName} ---\n`;
          fullText += XLSX.utils.sheet_to_txt(worksheet) + '\n\n';
        });
        text = fullText;
      } else if (extension === 'txt') {
        text = await file.text();
      } else {
        setError('Formato de arquivo não suportado. Tente PDF, DOCX, XLSX ou TXT.');
        setIsParsing(false);
        return;
      }

      if (!text.trim()) {
        throw new Error('O arquivo parece estar vazio ou não contém texto extraível.');
      }

      setInputText(text);
    } catch (err: any) {
      console.error('Erro ao ler arquivo:', err);
      setError(`Falha ao processar o arquivo: ${err.message || 'Erro desconhecido'}. Tente colar o texto manualmente.`);
    } finally {
      setIsParsing(false);
    }
  };

  const clearFile = () => {
    setFileName(null);
    setInputText('');
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleAudit = async () => {
    if (!inputText.trim()) {
      setError('Por favor, insira o conteúdo ou faça upload de um arquivo para análise.');
      return;
    }

    setIsLoading(true);
    setError(null);
    setAnalysis(null);
    setChatHistory([]);

    try {
      if (!process.env.GEMINI_API_KEY) {
        throw new Error('CONFIG_ERROR: Chave da API Gemini não detectada. Verifique as configurações de ambiente.');
      }

      // Senior Dev: Estabilizando para gemini-3.1-pro-preview que é o padrão atual desta infra
      const response = await genAI.models.generateContent({ 
        model: "gemini-3.1-pro-preview",
        contents: `Efetue a ${auditType} perante a norma ${selectedStandard} operando sobre este conteúdo técnico:\n\n${inputText}`,
        config: {
          systemInstruction: `
            CONTEXTO: AUDITOR LÍDER SÊNIOR (25+ ANOS DE EXPERIÊNCIA)
            OBJETIVO: Realizar auditoria técnico-normativa implacável no SiAC 2021 (PBQP-H) e ISO 9001:2015.
            TIPO DE AUDITORIA ATUAL: ${auditType}
            ESCOPO NORMATIVO: ${selectedStandard}

            🔥 DIRETRIZES DE RIGOR (MÉTODO UNITÀ):
            1. REQUISITO 4.1 e 4.2: Denuncie confusão entre cenário e partes interessadas.
            2. REQUISITO 7.1.5: Exija rastreabilidade metrológica absoluta.
            3. REQUISITO 8.4: Verifique rigorosamente o controle de serviços providos externamente.
            4. TERMINOLOGIA: Elimine adjetivos ("bom", "adequado"). Substitua por critérios quantitativos e normatizados.

            ESTRUTURA OBRIGATÓRIA DO RELATÓRIO:
            # 🔍 PARECER TÉCNICO DE AUDITORIA: ${selectedStandard}
            
            ## 💀 VEREDITO DE CERTIFICAÇÃO
            **[APROVADO | REPROVADO | CRÍTICO]**
            Justificativa estratégica baseada nos riscos de conformidade.

            ## 1. 🚨 NÃO-CONFORMIDADES DE MAIOR (BLOQUEANTES)
            Erros técnicos que impedem a certificação imediata.

            ## 2. 📉 OBSERVAÇÕES E GAPs DE "NÍVEL A"
            Falhas de detalhamento que impedem a excelência técnica.

            ## 3. 🛡️ ANÁLISE DE RISCO OPERACIONAL
            Consequências práticas em obra (retrabalho, patologias, acidentes).

            ## 4. ✍️ REESCRITA TÉCNICA SUGERIDA
            Parágrafos corrigidos com linguagem técnico-normativa.
          `,
          temperature: 0.2,
        },
      });

      const text = response.text;
      
      if (!text) throw new Error('EMPTY_RESPONSE: A IA não retornou conteúdo. Tente reduzir o escopo do texto.');
      
      setAnalysis(text);
      setConfidence(92); 
    } catch (err: any) {
      console.error('Expert Audit Error:', err);
      const errorMsg = err.message || 'Erro de processamento na nuvem.';
      setError(`Falha na análise: ${errorMsg}. Se o erro persistir, divida o documento em partes menores.`);
    } finally {
      setIsLoading(false);
    }
  };

  const handleDownloadPDF = async () => {
    if (!reportRef.current || isExporting) return;
    
    setIsExporting(true);
    setError(null);

    try {
      const element = reportRef.current;
      if (!element) throw new Error('ELEMENT_NOT_FOUND: O relatório não está pronto para exportação.');

      const opt: any = {
        margin: [10, 10, 10, 10],
        filename: `Unita-Parecer-${fileName?.split('.')[0] || 'Doc'}-${new Date().getTime()}.pdf`,
        image: { type: 'jpeg', quality: 0.95 },
        html2canvas: { 
          scale: 2, 
          useCORS: true,
          logging: false,
          letterRendering: true,
          allowTaint: true,
          onclone: (clonedDoc: Document) => {
            // FIX SÊNIOR: PURGE DE CORES MODERNAS (Tailwind 4)
            // O html2canvas falha ao processar oklch() e oklab().
            // Forçamos um fallback para HEX/RGB em todo o documento clonado.
            const style = clonedDoc.createElement('style');
            style.innerHTML = `
              :root {
                --color-slate-900: #0f172a !important;
                --color-slate-600: #1e293b !important;
                --color-blue-600: #2563eb !important;
              }
              * { 
                color-scheme: light !important;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
                border-color: #e2e8f0 !important;
              }
              body { background-color: #ffffff !important; }
              /* Override global para qualquer elemento que use as variáveis v4 */
              h1, h2, h3, h4, h5, h6, .text-slate-900 { color: #0f172a !important; }
              p, li, span, div, .text-slate-600, .text-slate-700 { color: #334155 !important; }
              strong { color: #020617 !important; }
              
              .bg-blue-600 { background-color: #2563eb !important; }
              .text-blue-600 { color: #2563eb !important; }
              .bg-white { background-color: #ffffff !important; }
              .bg-slate-50 { background-color: #f8fafc !important; }
              
              .border, .border-slate-200, .border-slate-100 { border-color: #e2e8f0 !important; }
              
              /* Mata qualquer gradiente que possa carregar cores oklch */
              .bg-gradient-to-r, .bg-gradient-to-b, .bg-gradient-to-br { 
                background-image: none !important; 
                background-color: #2563eb !important; 
              }
              
              /* Prevenção de quebra de layout no clone */
              .report-container { width: 100% !important; margin: 0 !important; padding: 20px !important; }
            `;
            clonedDoc.head.appendChild(style);
          }
        },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
        pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
      };

      await html2pdf().set(opt).from(element).save();
    } catch (err: any) {
      console.error('PDF Export Audit Error:', err);
      setError(`Falha técnica na exportação: ${err.message || 'Erro de Renderização'}. Tente Imprimir (Ctrl+P).`);
    } finally {
      setIsExporting(false);
    }
  };

  const handlePrint = () => {
    window.print();
  };

  const handleChat = async () => {
    if (!userQuestion.trim() || isChatting) return;

    const currentQuestion = userQuestion;
    setUserQuestion('');
    setChatHistory(prev => [...prev, { role: 'user', content: currentQuestion }]);
    setIsChatting(true);

    try {
      // Senior Dev: Usando o padrão correto do @google/genai
      const response = await genAI.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: `
          Contexto do Documento Analisado:
          "${inputText.substring(0, 50000)}" (Contexto do documento principal)

          Diagnóstico Gerado Anteriormente:
          "${analysis}"

          Pergunta Específica do Usuário:
          "${currentQuestion}"

          Responda como um Auditor Líder Especialista. Seja técnico, direto e aponte exatamente as falhas ou conformidades no documento com base no SiAC 2021/ISO 9001.
        `,
        config: {
          temperature: 0.3,
        }
      });
      
      setChatHistory(prev => [...prev, { role: 'assistant', content: response.text || 'Desculpe, não consegui processar a resposta.' }]);
    } catch (err) {
      console.error(err);
      setChatHistory(prev => [...prev, { role: 'assistant', content: 'Erro ao processar consulta. Verifique o limite de caracteres.' }]);
    } finally {
      setIsChatting(false);
      setTimeout(() => chatEndRef.current?.scrollIntoView({ behavior: 'smooth' }), 100);
    }
  };

  return (
    <div className="min-h-screen bg-[#F8FAFC]">
      {/* Topbar Corporativo */}
      <nav className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-6 sticky top-0 z-50 shadow-sm">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 p-2 rounded-lg">
            <ShieldCheck className="text-white" size={24} />
          </div>
          <div>
            <h1 className="text-lg font-black tracking-tight text-slate-900 leading-none">UNITÀ ENGENHARIA</h1>
            <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mt-1">Audit Platform • ISO 9001 / SiAC</p>
          </div>
        </div>

        <div className="hidden md:flex items-center gap-8">
          <div className="flex gap-1 text-[11px] font-bold uppercase tracking-wider text-slate-400">
            <span>Ambiente:</span>
            <span className="text-green-600">Produção</span>
          </div>
          <div className="flex gap-1 text-[11px] font-bold uppercase tracking-wider text-slate-400">
            <span>Base Normativa:</span>
            <span className="text-blue-600">Atualizada 2024</span>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <button 
            onClick={() => window.location.reload()}
            className="text-[10px] font-black uppercase tracking-widest px-4 py-2 bg-slate-100 text-slate-600 rounded-lg hover:bg-slate-200 transition-all"
          >
            Nova Análise
          </button>
        </div>
      </nav>

      <main className="flex h-[calc(100vh-64px)] overflow-hidden">
        {/* Sidebar de Preparação */}
        <aside className="w-80 bg-white border-r border-slate-200 overflow-y-auto hidden lg:flex flex-col">
          <div className="p-6 border-b border-slate-100">
            <h2 className="text-xs font-black text-slate-400 uppercase tracking-[0.2em] mb-4">Preparação da Auditoria</h2>
            
            <div className="space-y-6">
              {/* Documento Section */}
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-100">
                <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Documento Principal</p>
                {fileName ? (
                  <div className="flex items-center gap-3 overflow-hidden">
                    <div className="bg-blue-100 p-2 rounded-lg text-blue-600 shrink-0">
                      <FileText size={18} />
                    </div>
                    <span className="text-sm font-bold text-slate-700 truncate">{fileName}</span>
                  </div>
                ) : (
                  <button 
                    onClick={() => fileInputRef.current?.click()}
                    className="w-full py-4 border-2 border-dashed border-slate-200 rounded-xl flex flex-col items-center gap-2 hover:bg-blue-50 hover:border-blue-200 transition-all group"
                  >
                    <FileUp size={20} className="text-slate-400 group-hover:text-blue-600 transition-colors" />
                    <span className="text-[10px] font-black uppercase tracking-widest text-slate-400 group-hover:text-blue-600">Carregar Arquivo</span>
                  </button>
                )}
                <input 
                  type="file" 
                  ref={fileInputRef} 
                  onChange={handleFileUpload} 
                  className="hidden" 
                  accept=".pdf,.docx,.doc,.xlsx,.xls" 
                />
              </div>

              {/* Escopo Normativo */}
              <div>
                <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-3">Escopo Normativo</p>
                <div className="space-y-2">
                  <label className="flex items-center gap-3 p-3 rounded-lg hover:bg-slate-50 cursor-pointer transition-colors border border-transparent hover:border-slate-100">
                    <input type="checkbox" checked readOnly className="w-4 h-4 rounded border-slate-300 text-blue-600 focus:ring-blue-600" />
                    <span className="text-xs font-bold text-slate-700">ISO 9001:2015</span>
                  </label>
                  <label className="flex items-center gap-3 p-3 rounded-lg hover:bg-slate-50 cursor-pointer transition-colors border border-transparent hover:border-slate-100">
                    <input type="checkbox" checked readOnly className="w-4 h-4 rounded border-slate-300 text-blue-600 focus:ring-blue-600" />
                    <span className="text-xs font-bold text-slate-700">SiAC / PBQP-H (Nível A)</span>
                  </label>
                  <label className="flex items-center gap-3 p-3 rounded-lg opacity-50 cursor-not-allowed">
                    <input type="checkbox" disabled className="w-4 h-4 rounded border-slate-300" />
                    <span className="text-xs font-bold text-slate-400 text-decoration-line-through">Regras Internas (Manual Unità)</span>
                  </label>
                </div>
              </div>

              {/* Tipo de Auditoria */}
              <div>
                <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-3">Tipo de Auditoria</p>
                <div className="space-y-2">
                  {[
                    { id: 'diagnostic', label: 'Auditoria Diagnóstica' },
                    { id: 'internal', label: 'Auditoria Interna' },
                    { id: 'pre-cert', label: 'Pré-Auditoria de Certificação' }
                  ].map((type) => (
                    <label key={type.id} className={`flex items-center gap-3 p-3 rounded-lg cursor-pointer transition-all border ${auditType === type.id ? 'bg-blue-50 border-blue-100 text-blue-700' : 'hover:bg-slate-50 border-transparent text-slate-600'}`}>
                      <input 
                        type="radio" 
                        name="auditType" 
                        checked={auditType === type.id} 
                        onChange={() => setAuditType(type.id as any)}
                        className="w-4 h-4 border-slate-300 text-blue-600 focus:ring-blue-600" 
                      />
                      <span className="text-xs font-bold">{type.label}</span>
                    </label>
                  ))}
                </div>
              </div>

              {/* Botão de Ação Principal */}
              {fileName && !analysis && (
                <button 
                  onClick={handleAudit}
                  disabled={isLoading}
                  className="w-full py-4 bg-blue-600 text-white rounded-xl font-black uppercase tracking-widest text-[11px] shadow-lg shadow-blue-100 hover:bg-blue-700 active:scale-95 transition-all flex items-center justify-center gap-3 disabled:opacity-50"
                >
                  {isLoading ? <Loader2 size={18} className="animate-spin" /> : <Search size={18} />}
                  Executar Auditoria Técnica
                </button>
              )}
            </div>
          </div>

          <div className="mt-auto p-6 bg-slate-50 border-t border-slate-200">
            <div className="flex items-center gap-3 mb-2">
              <div className="h-2 w-2 rounded-full bg-green-500 animate-pulse" />
              <p className="text-[10px] font-black text-slate-600 uppercase tracking-widest">Sistema Operacional</p>
            </div>
            <p className="text-[10px] text-slate-400 font-bold leading-relaxed">
              Agente de IA configurado para rigor Nível A conforme Regimento SiAC 2021.
            </p>
          </div>
        </aside>

        {/* Dashboard Central / Área de Relatório */}
        <div className="flex-1 flex flex-col min-w-0 bg-slate-50 relative">
          <AnimatePresence>
            {!analysis && !isLoading && !isParsing && (
              <motion.div 
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95 }}
                className="absolute inset-0 flex flex-col items-center justify-center p-8 z-10 text-center"
              >
                <div className="max-w-md">
                  <div className="w-24 h-24 bg-blue-50 rounded-3xl flex items-center justify-center mb-8 mx-auto shadow-sm">
                    <Building2 className="text-blue-600" size={48} />
                  </div>
                  <h2 className="text-3xl font-black text-slate-900 mb-4 tracking-tight">Estação do Auditor</h2>
                  <p className="text-slate-500 text-lg mb-10 leading-relaxed font-medium">Carregue o manual, procedimento ou evidência para iniciar a análise técnico-normativa assistida por inteligência artificial.</p>
                  
                  <button 
                    onClick={() => fileInputRef.current?.click()}
                    className="px-10 py-5 bg-white border border-slate-200 text-slate-700 rounded-2xl font-black uppercase tracking-widest text-xs hover:border-blue-400 hover:text-blue-600 shadow-sm transition-all flex items-center gap-3 mx-auto"
                  >
                    <FileUp size={20} />
                    Selecionar Documento Principal
                  </button>
                </div>
              </motion.div>
            )}
          </AnimatePresence>

          {(isLoading || isParsing) && (
            <div className="absolute inset-0 flex items-center justify-center z-20 bg-slate-50/80 backdrop-blur-sm">
              <div className="text-center space-y-6">
                <div className="relative">
                  <div className="absolute inset-0 bg-blue-100 rounded-full animate-ping opacity-20" />
                  <div className="relative bg-white p-8 rounded-3xl shadow-xl shadow-blue-100 border border-blue-50 border-none">
                    <Loader2 className="animate-spin text-blue-600 mx-auto" size={48} />
                  </div>
                </div>
                <div>
                  <p className="text-xs font-black text-slate-400 uppercase tracking-[0.3em] mb-2">{isParsing ? 'Extração de Dados' : 'Análise Normativa'}</p>
                  <p className="text-slate-600 font-bold text-lg">{parsingProgress || 'Processando inteligência...'}</p>
                </div>
              </div>
            </div>
          )}

          {analysis && (
            <div className="flex-1 overflow-y-auto pt-8 px-6 md:px-12 pb-24 h-full scroll-smooth">
              <div className="max-w-5xl mx-auto space-y-8">
                {/* Header do Relatório */}
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 pb-8 border-b border-slate-200">
                  <div className="space-y-1">
                    <div className="flex items-center gap-2 text-slate-400 font-black uppercase tracking-widest text-[10px] mb-2">
                      <ClipboardCheck size={14} />
                      Auditoria Concluída
                    </div>
                    <h2 className="text-4xl font-black text-slate-900 tracking-tight">Parecer Técnico de Auditoria</h2>
                    <p className="text-slate-500 font-bold text-lg">Documento: {fileName}</p>
                  </div>
                  
                  <div className="flex items-center gap-3">
                    <button 
                      onClick={handlePrint}
                      className="p-3 bg-white text-slate-600 hover:text-blue-600 hover:border-blue-200 rounded-2xl transition-all duration-300 border border-slate-200 flex items-center gap-2 shadow-sm font-black uppercase tracking-widest text-[10px]"
                    >
                      <Printer size={18} />
                      <span className="hidden sm:inline">Imprimir</span>
                    </button>
                    
                    <button 
                      onClick={handleDownloadPDF}
                      disabled={isExporting}
                      className="p-3 bg-blue-600 text-white hover:bg-blue-700 rounded-2xl transition-all duration-300 shadow-lg shadow-blue-100 border border-blue-500 flex items-center gap-2 disabled:opacity-50 font-black uppercase tracking-widest text-[10px]"
                    >
                      {isExporting ? <Loader2 size={18} className="animate-spin" /> : <Download size={18} />}
                      <span className="hidden sm:inline">Salvar em PDF</span>
                    </button>
                  </div>
                </div>

                {/* Conteúdo do Relatório */}
                <div 
                  ref={reportRef} 
                  className="print-content bg-white p-12 md:p-20 rounded-[40px] shadow-xl shadow-slate-200/50 border border-slate-100 prose prose-slate prose-blue max-w-none prose-headings:font-black prose-h1:text-4xl prose-h1:tracking-tight prose-h1:mb-10 prose-h2:text-2xl prose-h2:mt-12 prose-h2:mb-6 prose-p:text-slate-600 prose-p:leading-relaxed prose-p:text-lg prose-li:text-slate-600 prose-li:text-lg prose-strong:text-slate-900 border-none outline-none"
                >
                  <ReactMarkdown>{analysis}</ReactMarkdown>
                  
                  {/* Chat Interface Integrada no Relatório */}
                  <div className="chat-interface mt-24 pt-16 border-t border-slate-100 not-prose">
                    <div className="flex items-center gap-4 mb-8">
                      <div className="bg-blue-600 p-3 rounded-2xl text-white shadow-lg shadow-blue-100">
                        <MessageSquare size={24} />
                      </div>
                      <div>
                        <h3 className="text-xl font-black text-slate-900 tracking-tight">Audit Deep-Dive</h3>
                        <p className="text-sm font-bold text-slate-400 uppercase tracking-widest">Tire dúvidas específicas sobre este documento</p>
                      </div>
                    </div>

                    <div className="space-y-6 mb-10">
                      {chatHistory.map((chat, index) => (
                        <div key={index} className={`flex ${chat.role === 'user' ? 'justify-end' : 'justify-start'}`}>
                          <div className={`max-w-[85%] p-6 rounded-3xl text-lg leading-relaxed ${
                            chat.role === 'user' 
                              ? 'bg-blue-600 text-white font-medium rounded-tr-none shadow-lg shadow-blue-100' 
                              : 'bg-slate-50 text-slate-700 border border-slate-100 rounded-tl-none font-medium'
                          }`}>
                            <ReactMarkdown>{chat.content}</ReactMarkdown>
                          </div>
                        </div>
                      ))}
                      {isChatting && (
                        <div className="flex justify-start">
                          <div className="bg-slate-50 p-6 rounded-3xl animate-pulse flex items-center gap-2">
                            <Loader2 size={24} className="animate-spin text-blue-600" />
                            <span className="text-slate-400 font-bold uppercase tracking-widest text-xs">O Auditor está processando...</span>
                          </div>
                        </div>
                      )}
                      <div ref={chatEndRef} />
                    </div>

                    <div className="relative group">
                      <input 
                        type="text" 
                        value={userQuestion}
                        onChange={(e) => setUserQuestion(e.target.value)}
                        onKeyPress={(e) => e.key === 'Enter' && handleChat()}
                        placeholder="Ex: Qual o item do SiAC fala sobre controle tecnológico?"
                        className="w-full bg-white border-2 border-slate-100 rounded-[32px] pl-8 pr-16 py-6 text-lg font-medium focus:ring-4 focus:ring-blue-100 focus:border-blue-400 outline-none transition-all shadow-sm"
                      />
                      <button 
                        onClick={handleChat}
                        disabled={isChatting || !userQuestion.trim()}
                        className="absolute right-3 top-3 bottom-3 w-20 bg-blue-600 text-white rounded-[24px] flex items-center justify-center hover:bg-blue-700 transition-all shadow-lg shadow-blue-100 disabled:opacity-50"
                      >
                        <Send size={24} />
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Feedback de Erro */}
          {error && (
            <div className="fixed bottom-8 left-1/2 -translate-x-1/2 w-full max-w-lg z-50">
              <motion.div 
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                className="bg-red-50 border border-red-100 p-6 rounded-3xl shadow-2xl flex items-start gap-4"
              >
                <div className="bg-red-100 p-2 rounded-xl text-red-600">
                  <X size={20} />
                </div>
                <div className="flex-1 text-red-900 font-medium leading-relaxed">
                  <p className="font-black uppercase tracking-widest text-[10px] mb-1">Erro de Auditoria</p>
                  {error}
                </div>
                <button onClick={() => setError(null)} className="text-red-400 hover:text-red-600">
                  <X size={20} />
                </button>
              </motion.div>
            </div>
          )}
        </div>
      </main>

      <footer className="bg-white border-t border-slate-200 py-12 px-8">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row justify-between items-center gap-8">
          <div className="flex items-center gap-3">
            <div className="bg-slate-100 p-2 rounded-lg">
              <ShieldCheck className="text-slate-400" size={20} />
            </div>
            <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
              Unità Audit Intelligence • 2024
            </p>
          </div>
          
          <p className="text-slate-400 text-xs font-medium">Esta ferramenta utiliza inteligência artificial baseada no PBQP-H Regimento Geral 2021 (Manual Unitário).</p>
        </div>
      </footer>
    </div>
  );
}
