/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useRef, useMemo } from 'react';
import { GoogleGenAI, ThinkingLevel } from "@google/genai";
import ReactMarkdown from 'react-markdown';
import { 
  Send, 
  Loader2, 
  CheckCircle2, 
  AlertCircle, 
  Sparkles, 
  FileText, 
  History,
  ExternalLink,
  Download,
  Check,
  Copy,
  ClipboardCheck,
  X,
  Upload,
  Link as LinkIcon,
  Type
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { Document, Packer, Paragraph, TextRun, DeletedTextRun, InsertedTextRun } from 'docx';
import { saveAs } from 'file-saver';
import { diffChars } from 'diff';
import * as mammoth from 'mammoth';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Types ---

type InputMode = 'text' | 'word' | 'gdoc';

interface Suggestion {
  original: string;
  suggested: string;
  reason: string;
}

interface ProofreadResult {
  summary: string;
  suggestions: Suggestion[];
  overallAdvice: string;
  sources?: { title: string; uri: string }[];
}

// --- App Component ---

export default function App() {
  const [inputMode, setInputMode] = useState<InputMode>('text');
  const [text, setText] = useState('');
  const [gdocUrl, setGdocUrl] = useState('');
  const [googleTokens, setGoogleTokens] = useState<any>(null);
  const [isFetchingGDoc, setIsFetchingGDoc] = useState(false);
  const [isPickerLoading, setIsPickerLoading] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [result, setResult] = useState<ProofreadResult | null>(null);
  const [selectedIndices, setSelectedIndices] = useState<number[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [isCopied, setIsCopied] = useState(false);
  const resultRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!file.name.endsWith('.docx')) {
      setError('Wordファイル（.docx）を選択してください。');
      return;
    }

    try {
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.extractRawText({ arrayBuffer });
      setText(result.value);
      setError(null);
    } catch (err) {
      console.error("File extraction error:", err);
      setError('ファイルの読み込みに失敗しました。');
    }
  };

  // Listen for OAuth success message
  useMemo(() => {
    const handleMessage = (event: MessageEvent) => {
      const origin = event.origin;
      if (!origin.endsWith('.run.app') && !origin.includes('localhost')) {
        return;
      }
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        setGoogleTokens(event.data.tokens);
        setError(null);
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

  const handleGoogleConnect = async () => {
    try {
      const response = await fetch('/api/auth/google/url');
      if (!response.ok) {
        throw new Error('Failed to get auth URL');
      }
      const { url } = await response.json();
      
      const authWindow = window.open(
        url,
        'google_oauth_popup',
        'width=600,height=700'
      );

      if (!authWindow) {
        setError('ポップアップがブロックされました。ブラウザの設定でポップアップを許可してください。');
      }
    } catch (err) {
      console.error("Google OAuth error:", err);
      setError('Google認証の開始に失敗しました。');
    }
  };

  const handleFetchGDoc = async () => {
    if (!gdocUrl) {
      setError('GoogleドキュメントのURLまたはIDを入力してください。');
      return;
    }

    if (!googleTokens?.access_token) {
      setError('Google Driveに接続してください。');
      return;
    }

    // Extract doc ID from URL if necessary
    let docId = gdocUrl;
    // Standard URL format: /d/ID/edit
    const match = gdocUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match) {
      docId = match[1];
    } else {
      // Check for ID only format (no URL)
      const idOnlyMatch = gdocUrl.match(/^[a-zA-Z0-9-_]+$/);
      if (!idOnlyMatch) {
        setError('GoogleドキュメントのURLが正しくありません。');
        return;
      }
    }

    setIsFetchingGDoc(true);
    setError(null);

    try {
      const response = await fetch('/api/gdoc/fetch', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ docId, accessToken: googleTokens.access_token }),
      });

      if (!response.ok) {
        const data = await response.json();
        throw new Error(data.error || 'ドキュメントの取得に失敗しました。');
      }

      const { content } = await response.json();
      setText(content);
      setError(null);
    } catch (err: any) {
      console.error("GDoc fetch error:", err);
      setError(err.message || 'ドキュメントの取得に失敗しました。');
    } finally {
      setIsFetchingGDoc(false);
    }
  };

  const handleOpenPicker = () => {
    if (!googleTokens?.access_token) {
      setError('Google Driveに接続してください。');
      return;
    }

    const apiKey = import.meta.env.VITE_GOOGLE_API_KEY;
    const clientId = import.meta.env.VITE_GOOGLE_CLIENT_ID;

    console.log("API Key loaded:", apiKey ? `${apiKey.substring(0, 5)}...` : "MISSING");
    console.log("Client ID loaded:", clientId ? `${clientId.substring(0, 10)}...` : "MISSING");

    if (!apiKey || !clientId) {
      setError('Google API Key または Client ID が設定されていません。Settingsから設定してください。');
      return;
    }

    setIsPickerLoading(true);

    const loadPicker = () => {
      console.log("Attempting to load Google Picker...");
      if (!(window as any).gapi) {
        setError('Google API (gapi) の初期化に失敗しました。広告ブロック等の拡張機能が干渉している可能性があります。');
        setIsPickerLoading(false);
        return;
      }

      (window as any).gapi.load('picker', {
        callback: () => {
          console.log("Picker API loaded successfully.");
          try {
            if (!(window as any).google || !(window as any).google.picker) {
              throw new Error('Google Picker API オブジェクトが見つかりません。');
            }

            const view = new (window as any).google.picker.View((window as any).google.picker.ViewId.DOCS);
            view.setMimeTypes('application/vnd.google-apps.document,application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            
            // オリジンの取得（iframe内での動作を考慮）
            const origin = window.location.origin || (window.location.protocol + '//' + window.location.host);
            console.log("Using Picker Origin:", origin);

            const pickerBuilder = new (window as any).google.picker.PickerBuilder()
              .enableFeature((window as any).google.picker.Feature.NAV_HIDDEN)
              .enableFeature((window as any).google.picker.Feature.SUPPORT_DRIVES)
              .setOAuthToken(googleTokens.access_token)
              .setDeveloperKey(apiKey)
              .addView(view)
              .setOrigin(origin)
              .setCallback((data: any) => {
                if (data.action === (window as any).google.picker.Action.PICKED) {
                  const doc = data.docs[0];
                  setGdocUrl(doc.url);
                  handleFetchGDocWithId(doc.id);
                }
                if (data.action === (window as any).google.picker.Action.CANCEL || data.action === (window as any).google.picker.Action.PICKED) {
                  setIsPickerLoading(false);
                }
              });

            // プロジェクト番号の設定（Client IDの最初の数字部分）
            const projectNumber = clientId.split('-')[0];
            if (projectNumber && /^\d+$/.test(projectNumber)) {
              pickerBuilder.setAppId(projectNumber);
            }

            const picker = pickerBuilder.build();
            picker.setVisible(true);
            console.log("Picker visibility set to true.");
          } catch (err: any) {
            console.error("Picker build error:", err);
            setError(`選択画面の起動に失敗しました: ${err.message || '不明なエラー'}`);
            setIsPickerLoading(false);
          }
        },
        onerror: () => {
          console.error("GAPI load error");
          setError('Google Picker APIの読み込みに失敗しました。ネットワーク設定を確認してください。');
          setIsPickerLoading(false);
        }
      });
    };

    if (!(window as any).gapi) {
      const script = document.createElement('script');
      script.src = 'https://apis.google.com/js/api.js';
      script.onload = loadPicker;
      script.onerror = () => {
        setError('Google APIスクリプトの読み込みに失敗しました。');
        setIsPickerLoading(false);
      };
      document.body.appendChild(script);
    } else {
      loadPicker();
    }
  };

  const handleFetchGDocWithId = async (docId: string) => {
    setIsFetchingGDoc(true);
    setError(null);

    try {
      const response = await fetch('/api/gdoc/fetch', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ docId, accessToken: googleTokens.access_token }),
      });

      if (!response.ok) {
        const data = await response.json();
        throw new Error(data.error || 'ドキュメントの取得に失敗しました。');
      }

      const { content } = await response.json();
      setText(content);
      setError(null);
    } catch (err: any) {
      console.error("GDoc fetch error:", err);
      setError(err.message || 'ドキュメントの取得に失敗しました。');
    } finally {
      setIsFetchingGDoc(false);
    }
  };

  const handleAnalyze = async () => {
    let finalInputText = text;

    if (inputMode === 'gdoc' && gdocUrl) {
      // If it's a gdoc URL, we'll try to use urlContext in the AI prompt
      // But for the diff calculation later, we need the text.
      // So we'll ask the AI to first extract the text or use the URL context.
      // For simplicity in this demo, we'll suggest the user to paste the content 
      // or we'll try to fetch it if it's a public export link.
      if (!text.trim()) {
        setError('Googleドキュメントの内容を以下に貼り付けるか、公開URLを指定してください。');
        return;
      }
    }

    if (!finalInputText.trim()) {
      setError('原稿を入力またはアップロードしてください。');
      return;
    }

    setIsAnalyzing(true);
    setError(null);
    setResult(null);
    setSelectedIndices([]);

    try {
      const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
      
      const prompt = `
        あなたは「神山まるごと高専（Kamiyama Marugoto College of Design, Engineering and Entrepreneurship）」の広報・編集のエキスパートです。
        以下の「文言統制表」のルールを厳守し、提供された記事の下書きを添削してください。

        ### 文言統制表（守るべきルール）
        1. **学校名**: 
           - 基本的に「神山まるごと高専」を使用。正式名称は「神山まるごと高等専門学校」。
           - 英語表記: "Kamiyama Marugoto College of Design, Technology, and Entrepreneurship"
        2. **ミッション・ビジョン**:
           - ミッション: 「テクノロジー×デザインで人間の未来を変える学校」
           - ビジョン: 「βメンタリティ」
           - 育成像: 「モノをつくる力で、コトを起こす」人
        3. **施設名（大文字アルファベット厳守）**:
           - 校舎: 「OFFICE」
           - 1期寮: 「HOME」
           - 2期寮: 「ROOMS」
        4. **カリキュラムの3つの柱（順序固定）**:
           - 「テクノロジー×デザイン×起業家精神」
        5. **呼称**:
           - 在籍者: 「学生」（「生徒」はNG）
           - 教職員: 「スタッフ」（「先生」とする場合は「スタッフ（神山まるごと高専における教職員の名称）」と注釈を入れる）
        6. **給食コンセプト**:
           - 「地産地食日本一をめざす」（「地産地消」や「日本一地産地食」は間違い）
        7. **表記揺れ**:
           - 「モノづくり」「コトを起こし」の表記を使用。
           - 数字はすべて「半角」を使用。
        8. **パートナー・基金**:
           - 「スカラーシップパートナー」「リソースサポーター」「プログラムパートナー」
           - 「宮田昇始スタートアップ基金」
           - 基金額: 「110億円相当」
        9. **教職員名簿（正誤判定の基準）**:
            以下の氏名と肩書きの組み合わせを「正」として、間違いがあれば修正してください。
            - 寺田親弘: 理事長
            - 五十棲浩二: 校長
            - 鈴木敦子: 副校長
            - 舟津潤: 事務部長 / 経営管理チームリーダー
            - 村山ザミット海優: クリエイティブディレクター
            - 大西栄樹: 広報・学生募集リーダー
            - 河野愛美: パートナー連携チームリーダー
            - 田中義崇: パートナー連携 / 寮 チームディレクター
            - 西川遥香: 看護師 / 保健師
            - 春田麻里: デザイン・エンジニアリング学科 国語
            - 廣瀬智子: デザイン・エンジニアリング学科 英語
            - 新井啓太: デザイン・エンジニアリング学科 デザイン教育
            - 鈴木佑奈: デザイン・エンジニアリング学科 保健体育
            - 藤川瞭: デザイン・エンジニアリング学科 助教 社会
            - 鈴木知真: デザイン・エンジニアリング学科 テクノロジー教育 (博士（工学）)
            - 阪本恒平: デザイン・エンジニアリング学科 国語
            - 水田徹: デザイン・エンジニアリング学科 助教 数学
            - 本末英樹: デザイン・エンジニアリング学科 准教授 デザイン教育
            - 越後正志: デザイン・エンジニアリング学科 准教授 デザイン教育 (博士（美術）)
            - 入江英也: デザイン・エンジニアリング学科 准教授 アントレプレナーシップ教育 (MBA)
            - 光永文彦: デザイン・エンジニアリング学科 准教授
            - 竹迫良範: デザイン・エンジニアリング学科 教授 テクノロジー教育
            - 松永歩: デザイン・エンジニアリング学科 講師 社会 (博士（政策科学）)
            - 齋藤亮次: デザイン・エンジニアリング学科 社会・アントレプレナーシップ教育
            - 小林勇輔: デザイン・エンジニアリング学科 物理
            - 須藤順: デザイン・エンジニアリング学科 准教授 アントレプレナーシップ教育
            - 山本周: デザイン・エンジニアリング学科 講師 数学・テクノロジー教育
            - 松本修平: デザイン・エンジニアリング学科 准教授 アントレプレナーシップ教育
            - 叶俊信: デザイン・エンジニアリング学科 テクノロジー教育
            - Queena Xu: デザイン・エンジニアリング学科 英語
            - 川﨑克寛: 寮長
            - 佐々木美優: 寮チーム
            - 小笠原愛: 寮チーム
            - 島浦舞: 経営管理チーム
            - 岡田真衣: 経営管理チーム
            - 蔵本有紀: 入試・広報戦略チーム
            - 山野誠: 学務チーム モノラボ担当
            - 山地駿徹: 学務チーム モノラボ担当
            - 後藤涼介: 学務チーム
            - 桑原菜穂: パートナー連携チーム
            - 北村美樹: パートナー連携チーム
            - 中村一彰: 起業キャリア応援チーム
            - 松坂孝紀: 理事

        【添削のポイント】
        - 上記の文言統制表に違反している箇所は必ず修正案を出すこと。
        - 誤字脱字、文法の修正。
        - 読者がワクワクするような魅力的な表現への提案。
        - **重要**: 修正提案（suggestions）の "original" は、記事本文から**一字一句違わず正確に**抜き出してください。

        【出力形式】
        以下のJSON形式で回答してください。
        {
          "summary": "記事の全体的な評価（100文字程度）",
          "suggestions": [
            {
              "original": "元の文章の一部（一字一句正確に）",
              "suggested": "修正後の文章",
              "reason": "文言統制表のどのルールに基づいたか、またはアドバイス"
            }
          ],
          "overallAdvice": "今後の執筆に役立つ全体的なアドバイス（マークダウン形式）"
        }

        【記事本文】
        ${finalInputText}
      `;

      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: prompt,
        config: {
          responseMimeType: "application/json",
          thinkingConfig: { thinkingLevel: ThinkingLevel.LOW },
        },
      });

      let jsonStr = response.text?.trim() || "{}";
      
      // Remove markdown code blocks if present
      if (jsonStr.startsWith("```")) {
        jsonStr = jsonStr.replace(/^```json\n?/, "").replace(/\n?```$/, "");
      }
      
      const parsedResult = JSON.parse(jsonStr) as ProofreadResult;
      
      // Extract grounding sources if available
      const chunks = response.candidates?.[0]?.groundingMetadata?.groundingChunks;
      if (chunks) {
        parsedResult.sources = chunks
          .filter(c => c.web)
          .map(c => ({ title: c.web!.title || '参考リンク', uri: c.web!.uri }));
      }

      setResult(parsedResult);
      // Initially select all suggestions
      setSelectedIndices(parsedResult.suggestions.map((_, i) => i));
      
      // Scroll to result after a short delay to allow rendering
      setTimeout(() => {
        resultRef.current?.scrollIntoView({ behavior: 'smooth' });
      }, 100);

    } catch (err) {
      console.error("Analysis error:", err);
      setError("添削中にエラーが発生しました。もう一度お試しください。");
    } finally {
      setIsAnalyzing(false);
    }
  };

  const toggleSuggestion = (index: number) => {
    setSelectedIndices(prev => 
      prev.includes(index) ? prev.filter(i => i !== index) : [...prev, index]
    );
  };

  const finalAppliedText = useMemo(() => {
    if (!result || !text) return text;
    
    let currentText = text;
    const sortedSuggestions = [...result.suggestions]
      .map((s, i) => ({ ...s, index: i }))
      .filter(s => selectedIndices.includes(s.index))
      .sort((a, b) => b.original.length - a.original.length);

    for (const s of sortedSuggestions) {
      currentText = currentText.replace(s.original, s.suggested);
    }
    
    return currentText;
  }, [text, result, selectedIndices]);

  const exportResult = async () => {
    if (!result) return;

    if (inputMode === 'word' || inputMode === 'gdoc') {
      // Export as Word (Google Docs is also best served as .docx)
      const diff = diffChars(text, finalAppliedText);
      const paragraphs: Paragraph[] = [];
      let currentChildren: any[] = [];

      diff.forEach((part, index) => {
        const lines = part.value.split('\n');
        lines.forEach((line, i) => {
          if (i > 0) {
            paragraphs.push(new Paragraph({ children: currentChildren }));
            currentChildren = [];
          }
          
          if (line.length > 0) {
            if (part.added) {
              currentChildren.push(new InsertedTextRun({
                text: line,
                author: "神山まるごと高専 記事添削くん",
                date: new Date().toISOString(),
                id: index,
              }));
            } else if (part.removed) {
              currentChildren.push(new DeletedTextRun({
                text: line,
                author: "神山まるごと高専 記事添削くん",
                date: new Date().toISOString(),
                id: index,
              }));
            } else {
              currentChildren.push(new TextRun({
                text: line,
              }));
            }
          }
        });
      });
      paragraphs.push(new Paragraph({ children: currentChildren }));

      const adviceParagraphs = result.overallAdvice.split('\n').map(line => 
        new Paragraph({
          children: [new TextRun({ text: line })]
        })
      );

      const doc = new Document({
        sections: [{
          properties: {},
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "神山まるごと高専 記事添削結果",
                  bold: true,
                  size: 32,
                }),
              ],
            }),
            new Paragraph({ text: "" }),
            ...paragraphs,
            new Paragraph({ text: "" }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "--- 全体アドバイス ---",
                  bold: true,
                }),
              ],
            }),
            ...adviceParagraphs,
          ],
        }],
      });

      const blob = await Packer.toBlob(doc);
      const filename = inputMode === 'word' ? "神山まるごと高専_添削済み記事.docx" : "神山まるごと高専_Googleドキュメント添削.docx";
      saveAs(blob, filename);
    } else {
      // Export as Text with diff markers
      const diff = diffChars(text, finalAppliedText);
      let output = "【神山まるごと高専 記事添削結果 (テキスト形式)】\n\n";
      
      diff.forEach(part => {
        if (part.added) {
          output += `[追加: ${part.value}]`;
        } else if (part.removed) {
          output += `[削除: ${part.value}]`;
        } else {
          output += part.value;
        }
      });

      output += "\n\n--- 全体アドバイス ---\n";
      output += result.overallAdvice;

      const blob = new Blob([output], { type: 'text/plain;charset=utf-8' });
      saveAs(blob, "神山まるごと高専_添削済み記事.txt");
    }
  };

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(finalAppliedText);
      setIsCopied(true);
      setTimeout(() => setIsCopied(false), 2000);
    } catch (err) {
      console.error("Failed to copy:", err);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-5xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="bg-indigo-600 p-2 rounded-lg">
              <Sparkles className="w-5 h-5 text-white" />
            </div>
            <div>
              <h1 className="font-bold text-xl tracking-tight text-slate-900 leading-none">
                神山まるごと高専 <span className="text-indigo-600">記事添削くん</span>
              </h1>
              <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest mt-1">
                Official Word Control Table Integrated
              </p>
            </div>
          </div>
          <div className="hidden sm:flex items-center gap-4 text-sm font-medium text-slate-500">
            <span className="flex items-center gap-1"><CheckCircle2 className="w-4 h-4" /> 校正</span>
            <span className="flex items-center gap-1"><Sparkles className="w-4 h-4" /> 魅力向上</span>
            <span className="flex items-center gap-1"><FileText className="w-4 h-4" /> 事実確認</span>
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-4 py-8 space-y-8">
        {/* Input Mode Selection */}
        <div className="flex p-1 bg-slate-200 rounded-xl w-fit mx-auto">
          <button
            onClick={() => setInputMode('text')}
            className={cn(
              "flex items-center gap-2 px-6 py-2 rounded-lg text-sm font-bold transition-all",
              inputMode === 'text' ? "bg-white text-indigo-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
            )}
          >
            <Type className="w-4 h-4" />
            テキスト
          </button>
          <button
            onClick={() => setInputMode('word')}
            className={cn(
              "flex items-center gap-2 px-6 py-2 rounded-lg text-sm font-bold transition-all",
              inputMode === 'word' ? "bg-white text-indigo-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
            )}
          >
            <FileText className="w-4 h-4" />
            Wordファイル
          </button>
          <button
            onClick={() => setInputMode('gdoc')}
            className={cn(
              "flex items-center gap-2 px-6 py-2 rounded-lg text-sm font-bold transition-all",
              inputMode === 'gdoc' ? "bg-white text-indigo-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
            )}
          >
            <LinkIcon className="w-4 h-4" />
            Googleドキュメント
          </button>
        </div>

        {/* Input Section */}
        <section className="space-y-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <div className="w-8 h-8 bg-indigo-100 rounded-lg flex items-center justify-center">
                {inputMode === 'text' ? <Type className="w-5 h-5 text-indigo-600" /> : 
                 inputMode === 'word' ? <FileText className="w-5 h-5 text-indigo-600" /> : 
                 <LinkIcon className="w-5 h-5 text-indigo-600" />}
              </div>
              <h2 className="text-lg font-bold text-slate-800">
                {inputMode === 'text' ? '原稿を貼り付ける' : 
                 inputMode === 'word' ? 'Wordファイルをアップロード' : 
                 'Googleドキュメントを読み込む'}
              </h2>
            </div>
            <div className="flex items-center gap-4">
              {text && (
                <button 
                  onClick={() => { setText(''); setGdocUrl(''); }}
                  className="text-xs font-medium text-slate-400 hover:text-red-500 transition-colors"
                >
                  入力をクリア
                </button>
              )}
              <span className="text-xs font-mono bg-slate-200 px-2 py-1 rounded text-slate-600">
                {text.length.toLocaleString()} 文字
              </span>
            </div>
          </div>
          
          <div className="relative group bg-white rounded-2xl border-2 border-slate-200 shadow-sm focus-within:border-indigo-500 focus-within:ring-4 focus-within:ring-indigo-50 transition-all overflow-hidden">
            {inputMode === 'word' && !text && (
              <div 
                onClick={() => fileInputRef.current?.click()}
                className="w-full h-80 flex flex-col items-center justify-center gap-4 cursor-pointer hover:bg-slate-50 transition-colors"
              >
                <div className="w-16 h-16 bg-indigo-50 rounded-full flex items-center justify-center">
                  <Upload className="w-8 h-8 text-indigo-600" />
                </div>
                <div className="text-center">
                  <p className="font-bold text-slate-700">クリックしてWordファイルを選択</p>
                  <p className="text-sm text-slate-400">または、ここにファイルをドラッグ＆ドロップ</p>
                </div>
                <input 
                  type="file" 
                  ref={fileInputRef} 
                  onChange={handleFileUpload} 
                  className="hidden" 
                  accept=".docx"
                />
              </div>
            )}

            {inputMode === 'gdoc' && (
              <div className="p-6 space-y-4">
                {!googleTokens ? (
                  <div className="flex flex-col items-center justify-center py-10 gap-4 border-2 border-dashed border-slate-200 rounded-xl bg-slate-50">
                    <div className="w-16 h-16 bg-white rounded-full flex items-center justify-center shadow-sm">
                      <LinkIcon className="w-8 h-8 text-indigo-600" />
                    </div>
                    <div className="text-center">
                      <p className="font-bold text-slate-700">Google Driveに接続</p>
                      <p className="text-sm text-slate-400 mb-4">ドキュメントを直接読み込むには認証が必要です</p>
                      <button
                        onClick={handleGoogleConnect}
                        className="px-6 py-2 bg-white border border-slate-200 rounded-lg text-sm font-bold text-slate-700 hover:bg-slate-50 transition-all shadow-sm flex items-center gap-2 mx-auto"
                      >
                        <img src="https://www.google.com/favicon.ico" className="w-4 h-4" alt="Google" referrerPolicy="no-referrer" />
                        Googleアカウントで接続
                      </button>
                    </div>
                  </div>
                ) : (
                  <div className="space-y-4">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2 text-sm font-bold text-emerald-600">
                        <CheckCircle2 className="w-4 h-4" />
                        Google Drive 接続済み
                      </div>
                      <button 
                        onClick={() => setGoogleTokens(null)}
                        className="text-xs text-slate-400 hover:text-red-500"
                      >
                        切断する
                      </button>
                    </div>
                    <div className="flex gap-2">
                      <input
                        type="url"
                        placeholder="GoogleドキュメントのURLまたはIDを入力..."
                        value={gdocUrl}
                        onChange={(e) => setGdocUrl(e.target.value)}
                        className="flex-1 p-3 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none"
                      />
                      <button
                        onClick={handleOpenPicker}
                        disabled={isPickerLoading || isFetchingGDoc}
                        className="px-4 py-2 bg-white border border-slate-200 rounded-xl text-sm font-bold text-slate-700 hover:bg-slate-50 transition-all shadow-sm flex items-center gap-2"
                      >
                        {isPickerLoading ? <Loader2 className="w-4 h-4 animate-spin" /> : <LinkIcon className="w-4 h-4" />}
                        ファイルを選択
                      </button>
                      <button
                        onClick={handleFetchGDoc}
                        disabled={isFetchingGDoc || !gdocUrl}
                        className={cn(
                          "px-6 py-2 rounded-xl font-bold transition-all shadow-md flex items-center gap-2",
                          isFetchingGDoc || !gdocUrl
                            ? "bg-slate-100 text-slate-400 cursor-not-allowed"
                            : "bg-indigo-600 text-white hover:bg-indigo-700"
                        )}
                      >
                        {isFetchingGDoc ? <Loader2 className="w-4 h-4 animate-spin" /> : <Download className="w-4 h-4" />}
                        読み込む
                      </button>
                    </div>
                    <p className="text-xs text-slate-400">
                      ※ 読み込んだ内容は下のテキストエリアに表示されます。
                    </p>
                  </div>
                )}
              </div>
            )}

            {(inputMode === 'text' || text || inputMode === 'gdoc') && (
              <>
                <div className="absolute top-4 left-4 pointer-events-none opacity-20 group-focus-within:opacity-5 transition-opacity">
                  <FileText className="w-12 h-12 text-slate-400" />
                </div>
                <textarea
                  className="w-full h-80 p-6 bg-transparent border-none focus:ring-0 text-slate-700 leading-relaxed text-lg placeholder:text-slate-300 resize-none"
                  placeholder={inputMode === 'gdoc' ? "Googleドキュメントの内容をここに貼り付けてください..." : "ここに記事の原稿を貼り付けてください..."}
                  value={text}
                  onChange={(e) => setText(e.target.value)}
                />
              </>
            )}
            
            <div className="p-4 bg-slate-50 border-t border-slate-100 flex justify-end items-center gap-4">
              <p className="text-xs text-slate-400 hidden sm:block">
                ※ 記事の内容はAIの学習には利用されません
              </p>
              <button
                onClick={handleAnalyze}
                disabled={isAnalyzing || !text.trim()}
                className={cn(
                  "flex items-center gap-2 px-8 py-3 rounded-xl font-bold transition-all shadow-lg",
                  isAnalyzing || !text.trim()
                    ? "bg-slate-200 text-slate-400 cursor-not-allowed shadow-none"
                    : "bg-indigo-600 text-white hover:bg-indigo-700 hover:shadow-indigo-200 active:scale-95"
                )}
              >
                {isAnalyzing ? (
                  <>
                    <Loader2 className="w-5 h-5 animate-spin" />
                    分析中...
                  </>
                ) : (
                  <>
                    <Sparkles className="w-5 h-5" />
                    添削を開始する
                  </>
                )}
              </button>
            </div>
          </div>
        </section>

        {/* Error Message */}
        <AnimatePresence>
          {error && (
            <motion.div
              initial={{ opacity: 0, y: -10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="p-4 bg-red-50 border border-red-100 rounded-xl flex items-center gap-3 text-red-700"
            >
              <AlertCircle className="w-5 h-5 shrink-0" />
              <p className="text-sm font-medium">{error}</p>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Result Section */}
        <div ref={resultRef}>
          <AnimatePresence>
            {result && (
              <motion.div
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                className="space-y-8"
              >
                {/* Summary Card */}
                <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm flex items-center justify-between">
                  <div>
                    <h3 className="text-sm font-bold text-indigo-600 uppercase tracking-wider mb-2">全体評価</h3>
                    <p className="text-lg text-slate-700 leading-relaxed font-medium">
                      {result.summary}
                    </p>
                  </div>
                  <button
                    onClick={exportResult}
                    className="flex items-center gap-2 px-6 py-3 bg-slate-900 text-white rounded-xl font-bold hover:bg-slate-800 transition-all shadow-lg active:scale-95 shrink-0 ml-4"
                  >
                    <Download className="w-5 h-5" />
                    {inputMode === 'text' ? 'テキストで出力' : 
                     inputMode === 'word' ? 'Wordで出力 (修正履歴あり)' : 
                     'Word形式で出力 (Googleドキュメント用)'}
                  </button>
                </div>

                {/* Suggestions List */}
                <div className="space-y-4">
                  <div className="flex items-center justify-between px-1">
                    <h3 className="text-lg font-semibold flex items-center gap-2">
                      <History className="w-5 h-5 text-indigo-600" />
                      修正箇所の選択
                    </h3>
                    <p className="text-xs text-slate-400">
                      採用する修正案にチェックを入れてください
                    </p>
                  </div>
                  <div className="grid gap-4">
                    {result.suggestions.map((s, i) => {
                      const isSelected = selectedIndices.includes(i);
                      return (
                        <motion.div
                          key={i}
                          initial={{ opacity: 0, x: -10 }}
                          animate={{ opacity: 1, x: 0 }}
                          transition={{ delay: i * 0.1 }}
                          onClick={() => toggleSuggestion(i)}
                          className={cn(
                            "bg-white border-2 rounded-xl overflow-hidden shadow-sm hover:shadow-md transition-all cursor-pointer relative",
                            isSelected ? "border-indigo-500 ring-2 ring-indigo-50" : "border-slate-200 opacity-70 grayscale-[0.5]"
                          )}
                        >
                          <div className="absolute top-4 right-4 z-10">
                            <div className={cn(
                              "w-6 h-6 rounded-full flex items-center justify-center transition-all",
                              isSelected ? "bg-indigo-600 text-white" : "bg-slate-100 text-slate-300"
                            )}>
                              {isSelected ? <Check className="w-4 h-4" /> : <X className="w-4 h-4" />}
                            </div>
                          </div>

                          <div className="p-4 border-b border-slate-100 bg-slate-50/50">
                            <span className="text-xs font-bold text-slate-400 uppercase">元の文章</span>
                            <p className="mt-1 text-slate-600 line-through decoration-red-300/50 pr-8">{s.original}</p>
                          </div>
                          <div className="p-4 bg-white">
                            <span className="text-xs font-bold text-indigo-500 uppercase">修正案</span>
                            <p className="mt-1 text-slate-900 font-medium pr-8">{s.suggested}</p>
                          </div>
                          <div className="p-4 bg-indigo-50/30 border-t border-indigo-50">
                            <p className="text-sm text-indigo-700 flex gap-2">
                              <Sparkles className="w-4 h-4 shrink-0 mt-0.5" />
                              <span>{s.reason}</span>
                            </p>
                          </div>
                        </motion.div>
                      );
                    })}
                  </div>
                </div>

                {/* Preview Section */}
                <div className="bg-slate-900 rounded-2xl p-8 shadow-xl text-white">
                  <div className="flex items-center justify-between mb-6">
                    <h3 className="text-lg font-bold flex items-center gap-2">
                      <CheckCircle2 className="w-5 h-5 text-emerald-400" />
                      最終プレビュー
                    </h3>
                    <div className="flex items-center gap-3">
                      <span className="text-xs px-2 py-1 bg-slate-800 rounded text-slate-400">
                        選択した修正が反映されています
                      </span>
                      <button
                        onClick={handleCopy}
                        className={cn(
                          "flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-bold transition-all border",
                          isCopied 
                            ? "bg-emerald-500/20 border-emerald-500/50 text-emerald-400" 
                            : "bg-slate-800 border-slate-700 text-slate-300 hover:bg-slate-700 hover:text-white"
                        )}
                      >
                        {isCopied ? (
                          <>
                            <ClipboardCheck className="w-3.5 h-3.5" />
                            コピーしました！
                          </>
                        ) : (
                          <>
                            <Copy className="w-3.5 h-3.5" />
                            テキストをコピー
                          </>
                        )}
                      </button>
                    </div>
                  </div>
                  <div className="bg-slate-800/50 rounded-xl p-6 text-slate-300 leading-relaxed whitespace-pre-wrap min-h-[200px]">
                    {finalAppliedText}
                  </div>
                </div>

                {/* Overall Advice */}
                <div className="bg-white border border-slate-200 rounded-2xl p-8 shadow-sm">
                  <h3 className="text-lg font-semibold mb-4 flex items-center gap-2">
                    <Sparkles className="w-5 h-5 text-indigo-600" />
                    さらなるブラッシュアップのために
                  </h3>
                  <div className="markdown-body prose prose-indigo">
                    <ReactMarkdown>{result.overallAdvice}</ReactMarkdown>
                  </div>
                </div>

                {/* Sources */}
                {result.sources && result.sources.length > 0 && (
                  <div className="bg-slate-100 rounded-2xl p-6">
                    <h3 className="text-sm font-bold text-slate-500 uppercase tracking-wider mb-4">参考情報 (Fact Check)</h3>
                    <div className="flex flex-wrap gap-3">
                      {result.sources.map((source, i) => (
                        <a
                          key={i}
                          href={source.uri}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 rounded-full text-sm text-slate-600 hover:text-indigo-600 hover:border-indigo-200 transition-all shadow-sm"
                        >
                          <span className="truncate max-w-[200px]">{source.title}</span>
                          <ExternalLink className="w-3 h-3" />
                        </a>
                      ))}
                    </div>
                  </div>
                )}
              </motion.div>
            )}
          </AnimatePresence>
        </div>

        {/* Empty State */}
        {!result && !isAnalyzing && (
          <div className="py-20 text-center space-y-4">
            <div className="inline-flex items-center justify-center w-16 h-16 bg-indigo-50 rounded-full mb-2">
              <Sparkles className="w-8 h-8 text-indigo-600" />
            </div>
            <h3 className="text-xl font-bold text-slate-900">あなたの記事を磨き上げましょう</h3>
            <p className="text-slate-500 max-w-md mx-auto">
              神山まるごと高専の魅力をより正確に、より情熱的に伝えるためのAIパートナーです。
              下書きを入力して「添削を開始」をクリックしてください。
            </p>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="max-w-5xl mx-auto px-4 py-12 border-t border-slate-200 text-center">
        <p className="text-sm text-slate-400">
          &copy; 2026 神山まるごと高専 記事添削くん | Powered by Gemini 3 Flash
        </p>
      </footer>
    </div>
  );
}
