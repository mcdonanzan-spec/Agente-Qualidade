import { useState, useRef } from 'react';
import { GoogleGenAI } from "@google/genai";
import ReactMarkdown from 'react-markdown';
import * as pdfjsLib from 'pdfjs-dist';
import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import { jsPDF } from "jspdf";
import { toPng } from 'html-to-image';
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
  Printer
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

  const handleAnalyze = async () => {
    if (!inputText.trim()) {
      setError('Por favor, insira o conteúdo ou faça upload de um arquivo para análise.');
      return;
    }

    setIsLoading(true);
    setError(null);
    setAnalysis(null);
    setChatHistory([]); // Reset chat for new document

    try {
      const systemInstruction = `
        Você é um AUDITOR LÍDER SÊNIOR especialista em certificações ISO 9001 e SiAC 2021 (PBQP-H). 
        Seu olhar é clínico, técnico e focado em evitar passivos para a construtora.

        MISSÃO CRÍTICA:
        Detectar erros conceituais "invisíveis" e vícios de linguagem que geram não-conformidades graves.

        🚨 ALERTA DE RIGOR TÉCNICO - O ERRO DO "CONTEXTO" (Cláusulas 4.1 e 4.2):
        - ANÁLISE OBRIGATÓRIA: Verifique se o documento confunde "Questões Externas/Internas" (ISO 9001 item 4.1) com "Partes Interessadas" (item 4.2).
        - O ERRO: Se o texto listar "Clientes", "Fornecedores", "Investidores", "Sócios", "Órgãos Públicos" ou "Comunidade" como 'Questões Externas ou Internas', VOCÊ DEVE APONTAR ISSO COMO UMA FALHA CONCEITUAL GRAVE. 
        - A RAZÃO: Pessoas e instituições são Partes Interessadas (4.2). As Questões (4.1) são FATOS, CONDIÇÕES, TENDÊNCIAS ou CIRCUNSTÂNCIAS que afetam a organização (ex: "Instabilidade do Preço do Aço", "Escassez de mão de obra qualificada", "Mudanças na legislação municipal", "Avanço tecnológico concorrente").
        - EXPLICAÇÃO DIDÁTICA: Explique que "Clientes" não são uma 'questão', mas sim uma 'parte interessada' que gera requisitos. A 'questão' seria, por exemplo, a "Exigência de novos prazos pelos clientes".

        PONTOS DE AUDITORIA DE CAMPO:
        - O documento define critérios de aceitação numéricos e tolerâncias técnicas (NBRs)?
        - Há menção explícita a FVS (Ficha de Verificação de Serviço) e FVM (Ficha de Verificação de Material)?
        - O fluxo de inspeção prevê "Ponto de Parada" (Hold Point) obrigatório?
        - O conteúdo respeita a Qualificação de Fornecedores e Mão de Obra do SiAC 2021.

        ESTRUTURA DO RELATÓRIO:
        - # 🔍 DIAGNÓSTICO AUDITOR SÊNIOR: SiAC 2021
        - ## 📈 ÍNDICE DE MATURIDADE NORMATIVA (0-100%)
          *Justifique a nota com base na profundidade técnica ou superficialidade detectada.*
        - ## 1. ANÁLISE "NAS ENTRELINHAS" (GAPs OCULTOS)
          *Destaque aqui erros de classificação como a confusão entre Contexto e Partes Interessadas.*
        - ## 2. NÃO-CONFORMIDADES DIRETAS (SiAC 2021 / ISO 9001)
        - ## 3. LISTA DE EVIDÊNCIAS FALTANTES (CHECKLIST DE CAMPO)
        - ## 4. CONCLUSÃO E RECOMENDAÇÃO ESTRATÉGICA
      `;

      const response = await genAI.models.generateContent({
        model: "gemini-2.0-flash", 
        contents: [{ role: "user", parts: [{ text: `Realize uma auditoria técnica rigorosa perante o SiAC 2021 neste conteúdo:\n\n${inputText}` }] }],
        config: {
          systemInstruction: systemInstruction,
          temperature: 0.1,
        },
      });

      setAnalysis(response.text || 'Não foi possível gerar a análise.');
    } catch (err: any) {
      console.error(err);
      setError('Falha na análise técnica. Se o arquivo for extremamente longo (mais de 100 páginas), tente processar por partes ou verifique sua conexão.');
    } finally {
      setIsLoading(false);
    }
  };

  const handleDownloadPDF = async () => {
    if (!reportRef.current || isExporting) return;
    
    setIsExporting(true);
    try {
      const dataUrl = await toPng(reportRef.current, { 
        quality: 0.95,
        backgroundColor: '#ffffff',
        style: {
          borderRadius: '0',
          boxShadow: 'none'
        }
      });
      
      const pdf = new jsPDF('p', 'mm', 'a4');
      const imgProps = pdf.getImageProperties(dataUrl);
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
      
      pdf.addImage(dataUrl, 'PNG', 0, 0, pdfWidth, pdfHeight);
      pdf.save(`Relatorio-Auditoria-${fileName || 'Qualidade'}.pdf`);
    } catch (err) {
      console.error('Erro ao gerar PDF:', err);
      alert('Falha ao gerar o PDF. Verifique se o navegador tem permissão.');
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
      const chatPrompt = `
        Contexto do Documento Analisado:
        "${inputText.substring(0, 50000)}" (Contexto do documento principal)

        Diagnóstico Gerado Anteriormente:
        "${analysis}"

        Pergunta Específica do Usuário:
        "${currentQuestion}"

        Responda como um Auditor Líder Especialista. Seja técnico, direto e aponte exatamente as falhas ou conformidades no documento com base no SiAC 2021/ISO 9001.
      `;

      const response = await genAI.models.generateContent({
        model: "gemini-2.0-flash",
        contents: [{ role: "user", parts: [{ text: chatPrompt }] }],
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
    <div className="min-h-screen bg-[#F8FAFC] text-slate-800 font-sans pb-12">
      {/* Navbar */}
      <nav className="bg-white border-b border-slate-200 px-6 py-4 sticky top-0 z-20 shadow-sm">
        <div className="max-w-7xl mx-auto flex justify-between items-center">
          <div className="flex items-center gap-3">
            <div className="bg-blue-600 p-2 rounded-lg text-white shadow-lg shadow-blue-200">
              <HardHat size={24} />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight text-slate-900 leading-none">Agente de Qualidade Civil</h1>
              <p className="text-xs text-slate-500 font-medium mt-1 uppercase tracking-wider underline decoration-blue-200 decoration-2 underline-offset-4">Análise Dinâmica de Documentos</p>
            </div>
          </div>
          <div className="hidden md:flex items-center gap-4 text-sm font-medium text-slate-600">
            <span className="flex items-center gap-1.5 bg-emerald-50 text-emerald-700 px-3 py-1 rounded-full border border-emerald-100 shadow-sm"><ShieldCheck size={16} /> ISO 9001:2015</span>
            <span className="flex items-center gap-1.5 bg-blue-50 text-blue-700 px-3 py-1 rounded-full border border-blue-100 shadow-sm"><Building2 size={16} /> PBQP-H / SiAC</span>
          </div>
        </div>
      </nav>

      <main className="max-w-7xl mx-auto p-6 grid grid-cols-1 lg:grid-cols-12 gap-8 mt-4">
        {/* Left Column: Input */}
        <div className="lg:col-span-5 space-y-6">
          <section className="bg-white rounded-[2rem] border border-slate-200 p-8 shadow-xl shadow-slate-200/50">
            <div className="flex items-center justify-between mb-8">
              <div className="flex items-center gap-3">
                <div className="bg-blue-100 p-2 rounded-xl text-blue-600">
                  <FileUp size={24} />
                </div>
                <h2 className="text-xl font-extrabold text-slate-900">Documentação Fonte</h2>
              </div>
            </div>
            
            {/* Dropzone Area */}
            <div 
              onClick={() => fileInputRef.current?.click()}
              className={`group relative border-2 border-dashed rounded-3xl p-10 flex flex-col items-center justify-center cursor-pointer transition-all duration-300 ${
                isParsing 
                  ? 'bg-slate-50 border-blue-400' 
                  : fileName 
                    ? 'bg-blue-50/20 border-blue-200' 
                    : 'bg-slate-50/50 border-slate-200 hover:border-blue-500 hover:bg-blue-50/50'
              }`}
            >
              <input 
                type="file" 
                ref={fileInputRef}
                onChange={handleFileUpload}
                accept=".pdf,.docx,.xlsx,.xls,.csv,.txt"
                className="hidden" 
              />
              
              <AnimatePresence mode="wait">
                {isParsing ? (
                  <motion.div 
                    key="parsing"
                    initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}
                    className="flex flex-col items-center gap-4"
                  >
                    <Loader2 className="animate-spin text-blue-600" size={40} />
                    <p className="text-slate-600 font-bold uppercase tracking-widest text-xs">{parsingProgress || 'Extraindo dados...'}</p>
                  </motion.div>
                ) : fileName ? (
                  <motion.div 
                    key="file-ready"
                    initial={{ opacity: 0, scale: 0.9 }} animate={{ opacity: 1, scale: 1 }}
                    className="flex flex-col items-center gap-4 w-full"
                  >
                    <div className="bg-white p-4 rounded-2xl shadow-md text-blue-600 border border-blue-50">
                      <FileText size={40} />
                    </div>
                    <div className="text-center w-full">
                      <p className="text-slate-900 font-bold truncate max-w-full px-4">{fileName}</p>
                      <button 
                        onClick={(e) => { e.stopPropagation(); clearFile(); }}
                        className="text-[10px] font-black text-red-500 hover:text-red-600 mt-3 flex items-center gap-1 justify-center mx-auto uppercase tracking-tighter bg-red-50 px-3 py-1 rounded-full transition-colors"
                      >
                        <X size={12} /> Descartar Arquivo
                      </button>
                    </div>
                  </motion.div>
                ) : (
                  <motion.div 
                    key="idle"
                    initial={{ opacity: 0 }} animate={{ opacity: 1 }}
                    className="flex flex-col items-center"
                  >
                    <div className="bg-white p-5 rounded-full shadow-lg mb-4 text-slate-400 group-hover:text-blue-500 group-hover:scale-110 transition-all">
                      <Upload size={32} />
                    </div>
                    <p className="text-slate-700 font-extrabold text-sm">Carregar Diretriz Técnica</p>
                    <p className="text-slate-400 text-[10px] mt-2 font-medium uppercase tracking-widest">PDF • Word • Excel • TXT</p>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>

            {/* Manual Text Preview/Edit */}
            <div className="mt-8">
              <div className="flex items-center justify-between mb-3 px-1">
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em]">Conteúdo Identificado</label>
                {inputText && (
                  <div className="text-[10px] font-bold text-emerald-600 flex items-center gap-1">
                    <CheckCircle2 size={12} /> {inputText.length} CARACTERES
                  </div>
                )}
              </div>
              <textarea
                className="w-full h-44 p-5 bg-slate-50 border border-slate-200 rounded-3xl focus:ring-4 focus:ring-blue-100 focus:border-blue-400 outline-none transition-all text-xs resize-none font-mono text-slate-600 leading-relaxed shadow-inner"
                placeholder="O conteúdo do documento processado aparecerá aqui para ajustes manuais caso necessário..."
                value={inputText}
                onChange={(e) => setInputText(e.target.value)}
              />
            </div>

            <div className="mt-8 space-y-4">
              <label className="block text-[10px] font-black text-slate-400 uppercase tracking-[0.2em] px-1">Referência Normativa</label>
              <div className="grid grid-cols-3 gap-3">
                {(['ISO 9001', 'SiAC (PBQP-H)', 'Both'] as Standard[]).map((std) => (
                  <button
                    key={std}
                    onClick={() => setSelectedStandard(std)}
                    className={`px-3 py-3 text-[11px] rounded-2xl font-black transition-all border-2 ${
                      selectedStandard === std
                        ? 'bg-blue-600 text-white border-blue-600 shadow-xl shadow-blue-100 transform -translate-y-1'
                        : 'bg-white text-slate-500 border-slate-100 hover:border-slate-200'
                    }`}
                  >
                    {std === 'Both' ? 'GERAL' : std}
                  </button>
                ))}
              </div>
            </div>

            <button
              onClick={handleAnalyze}
              disabled={isLoading || !inputText || isParsing}
              className={`w-full mt-10 py-5 px-8 rounded-3xl font-black text-sm uppercase tracking-[0.2em] flex items-center justify-center gap-4 transition-all duration-300 ${
                isLoading || !inputText || isParsing
                  ? 'bg-slate-100 text-slate-300 cursor-not-allowed border border-slate-100'
                  : 'bg-slate-900 text-white hover:bg-blue-600 shadow-2xl shadow-blue-200 hover:shadow-blue-300 active:scale-[0.98]'
              }`}
            >
              {isLoading ? (
                <>
                  <Loader2 className="animate-spin" size={20} />
                  Auditoria AI Ativa...
                </>
              ) : (
                <>
                  <Search size={22} strokeWidth={3} />
                  Iniciar Auditoria AI
                </>
              )}
            </button>
            
            {error && (
              <motion.div 
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                className="mt-6 p-4 bg-rose-50 border border-rose-100 rounded-2xl flex items-start gap-3 text-rose-700 text-sm"
              >
                <AlertCircle size={20} className="shrink-0 mt-0.5" />
                <p className="font-bold leading-tight">{error}</p>
              </motion.div>
            )}
          </section>

          <div className="bg-gradient-to-br from-blue-700 to-indigo-800 rounded-[2rem] p-8 text-white shadow-2xl relative overflow-hidden group border border-blue-500/30">
            <div className="absolute -bottom-6 -right-6 p-4 opacity-10 transform -rotate-12 group-hover:scale-125 group-hover:-rotate-6 transition-all duration-700">
              <ShieldCheck size={180} />
            </div>
            <div className="relative z-10">
              <h3 className="font-black text-xl mb-3 flex items-center gap-2">
                Pilar da Conformidade
              </h3>
              <p className="text-sm text-blue-100/90 leading-relaxed font-medium">
                O Agente analisa não apenas o texto principal, mas procura por referências cruzadas entre diretrizes de obra e os requisitos mínimos do SiAC nível A e B.
              </p>
              <div className="mt-6 flex gap-2">
                <div className="bg-white/10 backdrop-blur-sm px-3 py-1 rounded-full text-[10px] font-bold uppercase tracking-widest border border-white/5">Precisão Técnica</div>
                <div className="bg-white/10 backdrop-blur-sm px-3 py-1 rounded-full text-[10px] font-bold uppercase tracking-widest border border-white/5">Foco SiAC</div>
              </div>
            </div>
          </div>
        </div>

        {/* Right Column: Results */}
        <div className="lg:col-span-7">
          <AnimatePresence mode="wait">
            {!analysis && !isLoading ? (
              <motion.div
                key="empty"
                initial={{ opacity: 0, y: 30 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -30 }}
                className="h-full flex flex-col items-center justify-center text-center p-16 border-2 border-dashed border-slate-200 rounded-[3rem] bg-white/40 backdrop-blur-sm"
              >
                <div className="bg-white p-8 rounded-[2.5rem] shadow-xl shadow-slate-100 mb-8 text-slate-200 ring-1 ring-slate-50">
                  <ClipboardCheck size={64} />
                </div>
                <h3 className="text-2xl font-black text-slate-900 uppercase tracking-tight">Relatório em Espera</h3>
                <p className="text-slate-400 text-sm max-w-sm mt-4 leading-relaxed font-medium">
                  Aguardando upload de documentos para iniciar a varredura normativa e gerar o diagnóstico técnico.
                </p>
                <div className="mt-10 grid grid-cols-2 gap-4 w-full max-w-sm">
                  <div className="bg-white p-4 rounded-3xl border border-slate-100 text-left">
                    <div className="w-8 h-8 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center mb-3">
                      <Search size={16} />
                    </div>
                    <p className="text-[10px] font-black text-slate-400 uppercase mb-1">Passo 1</p>
                    <p className="text-xs font-bold text-slate-700">Suba o PDF ou DOCX da qualidade</p>
                  </div>
                  <div className="bg-white p-4 rounded-3xl border border-slate-100 text-left">
                    <div className="w-8 h-8 bg-emerald-50 text-emerald-600 rounded-xl flex items-center justify-center mb-3">
                      <ShieldCheck size={16} />
                    </div>
                    <p className="text-[10px] font-black text-slate-400 uppercase mb-1">Passo 2</p>
                    <p className="text-xs font-bold text-slate-700">Receba a análise normativa</p>
                  </div>
                </div>
              </motion.div>
            ) : (
              <motion.div
                key="results"
                initial={{ opacity: 0, scale: 0.98 }}
                animate={{ opacity: 1, scale: 1 }}
                className="bg-white rounded-[2.5rem] border border-slate-200 shadow-2xl overflow-hidden min-h-[800px] flex flex-col group"
              >
                <div className="border-b border-slate-100 px-10 py-8 bg-white flex justify-between items-center relative overflow-hidden">
                  <div className="absolute top-0 right-0 h-1 w-full bg-gradient-to-r from-blue-600 via-indigo-500 to-emerald-400 opacity-80" />
                  <div className="flex items-center gap-4">
                    <div className="bg-emerald-500 p-2.5 rounded-2xl shadow-lg shadow-emerald-100">
                      <CheckCircle2 className="text-white" size={24} />
                    </div>
                    <div>
                      <span className="font-black text-slate-900 block text-xl tracking-tight uppercase">Dashboard do Auditor</span>
                      <div className="flex items-center gap-2 mt-1">
                        <span className="text-[9px] font-black text-slate-400 bg-slate-50 px-2 py-0.5 rounded border border-slate-100 uppercase tracking-widest leading-none">NORMA: {selectedStandard}</span>
                        <span className="text-[9px] font-black text-blue-600 bg-blue-50 px-2 py-0.5 rounded border border-blue-100 uppercase tracking-widest leading-none">Status: {isLoading ? 'Auditoria em curso' : 'Concluída'}</span>
                      </div>
                    </div>
                  </div>
                  <div className="flex gap-3">
                    <button 
                      onClick={handlePrint}
                      className="p-3 bg-slate-50 text-slate-600 hover:text-blue-600 hover:bg-blue-50 rounded-2xl transition-all duration-300 border border-slate-100 flex items-center gap-2 group/btn"
                      title="Imprimir Relatório"
                    >
                      <Printer size={18} strokeWidth={2.5} />
                      <span className="hidden sm:inline text-[10px] font-black uppercase tracking-widest">Imprimir</span>
                    </button>
                    
                    <button 
                      onClick={handleDownloadPDF}
                      disabled={isExporting}
                      className="p-3 bg-blue-600 text-white hover:bg-blue-700 rounded-2xl transition-all duration-300 shadow-lg shadow-blue-100 border border-blue-500 flex items-center gap-2 disabled:opacity-50"
                      title="Salvar como PDF"
                    >
                      {isExporting ? (
                        <Loader2 size={18} className="animate-spin" />
                      ) : (
                        <Download size={18} strokeWidth={2.5} />
                      )}
                      <span className="hidden sm:inline text-[10px] font-black uppercase tracking-widest">Salvar PDF</span>
                    </button>
                  </div>
                </div>

                <div ref={reportRef} className="print-content flex-1 p-10 md:p-16 overflow-y-auto prose prose-slate prose-blue max-w-none prose-headings:font-black prose-h1:text-4xl prose-h1:tracking-tight prose-h1:mb-10 prose-h2:text-2xl prose-h2:mt-12 prose-h2:mb-6 prose-p:text-slate-600 prose-p:leading-relaxed prose-p:text-lg prose-li:text-slate-600 prose-li:text-lg prose-strong:text-slate-900 border-none outline-none">
                  {isLoading ? (
                    <div className="flex flex-col items-center justify-center h-full space-y-8 py-32">
                      <div className="relative group">
                        <div className="w-24 h-24 border-4 border-slate-50 border-t-blue-600 rounded-full animate-spin transition-all duration-500 group-hover:scale-110" />
                        <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 bg-white p-3 rounded-2xl shadow-xl shadow-blue-100">
                          <HardHat className="text-blue-600" size={32} />
                        </div>
                      </div>
                      <div className="text-center space-y-2">
                        <p className="text-slate-900 font-black text-2xl uppercase tracking-tight">Varredura AI...</p>
                        <p className="text-slate-400 font-bold text-sm uppercase tracking-widest px-8">Processando cláusulas ocultas, riscos de obra e registros técnicos</p>
                      </div>
                    </div>
                  ) : (
                    <div className="space-y-12">
                      <div className="markdown-body animate-in fade-in slide-in-from-bottom-4 duration-1000 p-6 bg-slate-50/50 rounded-[2rem] border border-slate-100 shadow-inner">
                        <ReactMarkdown>{analysis || ''}</ReactMarkdown>
                      </div>

                      {/* Chat Interface */}
                      <div className="chat-interface mt-16 border-t pt-12">
                        <div className="flex items-center gap-3 mb-6">
                          <div className="bg-blue-600 p-2 rounded-xl text-white">
                            <ClipboardCheck size={20} />
                          </div>
                          <div>
                            <h3 className="font-black text-slate-900 text-lg">Consultoria Técnica Dinâmica</h3>
                            <p className="text-xs text-slate-400 font-bold uppercase tracking-wider">Tire dúvidas específicas sobre este documento</p>
                          </div>
                        </div>

                        <div className="space-y-4 mb-8 max-h-[400px] overflow-y-auto px-1">
                          {chatHistory.map((msg, idx) => (
                            <motion.div
                              key={idx}
                              initial={{ opacity: 0, x: msg.role === 'user' ? 20 : -20 }}
                              animate={{ opacity: 1, x: 0 }}
                              className={`flex ${msg.role === 'user' ? 'justify-end' : 'justify-start'}`}
                            >
                              <div className={`max-w-[85%] p-4 rounded-3xl text-sm font-medium shadow-sm ${
                                msg.role === 'user' 
                                  ? 'bg-blue-600 text-white rounded-tr-none' 
                                  : 'bg-white border border-slate-200 text-slate-700 rounded-tl-none'
                              }`}>
                                {msg.content}
                              </div>
                            </motion.div>
                          ))}
                          {isChatting && (
                            <div className="flex justify-start">
                              <div className="bg-slate-100 p-4 rounded-3xl rounded-tl-none flex gap-2">
                                <div className="w-2 h-2 bg-slate-300 rounded-full animate-bounce" />
                                <div className="w-2 h-2 bg-slate-300 rounded-full animate-bounce delay-75" />
                                <div className="w-2 h-2 bg-slate-300 rounded-full animate-bounce delay-150" />
                              </div>
                            </div>
                          )}
                          <div ref={chatEndRef} />
                        </div>

                        <div className="relative">
                          <input
                            type="text"
                            placeholder="Ex: Qual o item do SiAC fala sobre controle de rejunte?"
                            value={userQuestion}
                            onChange={(e) => setUserQuestion(e.target.value)}
                            onKeyDown={(e) => e.key === 'Enter' && handleChat()}
                            className="w-full bg-slate-50 border border-slate-200 rounded-2xl py-4 pl-6 pr-16 text-sm font-bold focus:ring-4 focus:ring-blue-100 focus:border-blue-400 outline-none transition-all shadow-inner"
                          />
                          <button
                            onClick={handleChat}
                            disabled={!userQuestion.trim() || isChatting}
                            className="absolute right-2 top-2 p-3 bg-blue-600 text-white rounded-xl hover:bg-blue-700 transition-all disabled:opacity-50"
                          >
                            <Search size={18} strokeWidth={3} />
                          </button>
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </main>

      <footer className="max-w-7xl mx-auto px-6 py-16 flex flex-col md:flex-row justify-between items-center gap-10 text-slate-400 text-xs border-t border-slate-200 mt-20">
        <div className="flex flex-col items-center md:items-start gap-3">
          <div className="flex items-center gap-2 text-slate-600">
            <HardHat size={20} strokeWidth={2.5} />
            <span className="font-black tracking-[0.3em] uppercase text-[10px]">Agente de Qualidade Civil AI</span>
          </div>
          <div className="flex gap-4">
             <span className="hover:text-blue-500 transition-colors cursor-help border-b border-transparent hover:border-blue-200">Documentação SiAC 2021</span>
             <span className="hover:text-blue-500 transition-colors cursor-help border-b border-transparent hover:border-blue-200">Suporte Técnico</span>
             <span className="hover:text-blue-500 transition-colors cursor-help border-b border-transparent hover:border-blue-200">Privacidade</span>
          </div>
        </div>
        <div className="text-center md:text-right space-y-1">
          <p className="font-black text-slate-500">© {new Date().getFullYear()} BUILDING QUALITY INTELLIGENCE</p>
          <p className="font-medium">Otimizado para Regimento SiAC 2021 e ISO 9001:2015</p>
          <p className="text-slate-300 font-medium">Esta ferramenta utiliza inteligência artificial baseada no PBQP-H Regimento Geral 2021.</p>
        </div>
      </footer>
    </div>
  );
}
