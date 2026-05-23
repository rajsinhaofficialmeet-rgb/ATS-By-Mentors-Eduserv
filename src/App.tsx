/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, {useState, useEffect, useRef} from 'react';
import * as pdfjsLib from 'pdfjs-dist';
import { GoogleGenAI, Type } from "@google/genai";
import { Copy, FileText, CheckCircle, XCircle, AlertCircle, Award, Download, Eye, X, ChevronRight, History, Star, LogOut, User } from "lucide-react";
import { motion, AnimatePresence } from "framer-motion";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType } from "docx";
import { saveAs } from "file-saver";
import { initializeApp } from 'firebase/app';
import { getAuth, signInAnonymously, GoogleAuthProvider, signInWithPopup, signOut } from 'firebase/auth';
import { getFirestore, addDoc, collection, getDocs, query, where } from 'firebase/firestore';
import firebaseConfig from '../firebase-applet-config.json';

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);
const auth = getAuth();
const googleProvider = new GoogleAuthProvider();

// IMPORTANT: Configure worker for pdfjs
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js`;

interface CandidateResult {
  name: string;
  email?: string;
  phone?: string;
  jobProfile?: string;
  filename: string;
  resumeUrl?: string;
  location?: string;
  score: number;
  yearsOfExperience: number;
  skills: string[];
  justification: string;
  verdict: 'Excellent' | 'Good' | 'Average' | 'Poor';
  education: string[];
  workExperience: string[];
  portfolioLink?: string;
}

interface EvaluationResults {
  candidates: CandidateResult[];
  insights: string;
  shortlistQuality: 'A+' | 'A' | 'B' | 'C';
  matchingCriteria: string[];
}

export default function App() {
  const [view, setView] = useState<'landing' | 'ats' | 'interview' | 'history' | 'shortlist'>('landing');
  const [mode, setMode] = useState<'ai' | 'simple' | 'import'>('ai');
  const [selectedCandidateFilenames, setSelectedCandidateFilenames] = useState<string[]>([]);
  const [shortlistedCandidateFilenames, setShortlistedCandidateFilenames] = useState<string[]>([]);
  const [showOnlyShortlisted, setShowOnlyShortlisted] = useState(false);
  const [jobRole, setJobRole] = useState('');
  const [jobDescription, setJobDescription] = useState('');
  const [resumes, setResumes] = useState<{file: File, status: 'Pending' | 'Parsing' | 'Parsed' | 'Error'}[]>([]);
  const [results, setResults] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [filterQuery, setFilterQuery] = useState('');
  const [filterVerdict, setFilterVerdict] = useState<string>('All');
  const [minScore, setMinScore] = useState<number>(0);
  const [selectedCandidate, setSelectedCandidate] = useState<CandidateResult | null>(null);
  const [history, setHistory] = useState<any[]>([]);
  const [recruiter, setRecruiter] = useState<{name: string, email: string, isAdmin: boolean, accessToken?: string, spreadsheetId?: string} | null>(null);
  const [theme, setTheme] = useState<'light' | 'dark'>('dark');
  const [loginMode, setLoginMode] = useState<'google' | 'recruiter' | null>(null);
  const [showLoginSelection, setShowLoginSelection] = useState(false);
  const [loginUsername, setLoginUsername] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [isAuthed, setIsAuthed] = useState(false);
  const [spreadsheetIdInput, setSpreadsheetIdInput] = useState('');
  const [resumeUrl, setResumeUrl] = useState<string | null>(null);

  // Email validation regex
  const validateEmail = (email: string) => /\S+@\S+\.\S+/.test(email);

  const handleProceedToInterview = async (candidate: CandidateResult) => {
    if (!recruiter?.accessToken) {
        alert("Authentication failed. Please login again.");
        return;
    }
    const SPREADSHEET_ID = recruiter?.spreadsheetId || spreadsheetIdInput;
    if (!SPREADSHEET_ID || SPREADSHEET_ID === "YOUR_SPREADSHEET_ID") {
        alert("Please set a valid Spreadsheet ID first.");
        return;
    }

    if (!candidate.email || !validateEmail(candidate.email)) {
        alert("Candidate does not have a valid email.");
        return;
    }

    if (!confirm(`Proceed to interview for ${candidate.name}?`)) return;

    try {
        // 1. Add to Sheet
        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/Sheet1!A1:append?valueInputOption=RAW`, {
            method: 'POST',
            headers: { 
                Authorization: `Bearer ${recruiter.accessToken}`, 
                'Content-Type': 'application/json' 
            },
            body: JSON.stringify({ values: [[candidate.name, candidate.email, candidate.jobProfile, candidate.location || 'N/A', new Date().toISOString()]] })
        });

        // 2. Send Email
        const emailBody = `Subject: Interview Invitation\n\nDear ${candidate.name},\n\nWe are pleased to invite you to an interview.\n\nBest regards,\n${recruiter.name}`;
        await fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages/send`, {
            method: 'POST',
            headers: { 
                Authorization: `Bearer ${recruiter.accessToken}`, 
                'Content-Type': 'application/json' 
            },
            body: JSON.stringify({ raw: btoa(emailBody) })
        });
        
        alert("Candidate moved to interview and email sent!");
    } catch (e) {
        console.error(e);
        alert("Failed to proceed/send email.");
    }
  };

  const handleExportShortlist = async () => {
    if (!recruiter?.accessToken) {
        alert("Authentication failed. Please login again.");
        return;
    }
    const SPREADSHEET_ID = recruiter?.spreadsheetId || spreadsheetIdInput;
    if (!SPREADSHEET_ID || SPREADSHEET_ID === "YOUR_SPREADSHEET_ID") {
        alert("Please set a valid Spreadsheet ID first.");
        return;
    }
    
    const candidatesToExport = evalResults?.candidates.filter(c => shortlistedCandidateFilenames.includes(c.filename));
    
    if (!candidatesToExport || candidatesToExport.length === 0) {
        alert("No shortlisted candidates to export.");
        return;
    }

    try {
        const response = await fetch('/api/sheets/export', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ 
                spreadsheetId: SPREADSHEET_ID, 
                accessToken: recruiter.accessToken, 
                candidates: candidatesToExport 
            })
        });

        if (!response.ok) throw new Error('Failed to export');
        alert(`Exported ${candidatesToExport.length} candidates to spreadsheet!`);
    } catch (e) {
        console.error(e);
        alert("Failed to export candidates.");
    }
  };

  const handleLogout = async () => {
    try {
      await signOut(auth);
      setRecruiter(null);
      setIsAuthed(false);
      setView('landing');
    } catch (e) {
      console.error(e);
      alert("Logout failed");
    }
  };

  const handleCustomLogin = async () => {
    try {
      const response = await fetch('/api/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: loginUsername, password: loginPassword })
      });
      if (!response.ok) throw new Error('Login failed');
      const data = await response.json();
      setRecruiter({
        name: data.user.name,
        email: data.user.email,
        isAdmin: data.user.isAdmin
      });
      setIsAuthed(true);
      setView('ats');
      fetchHistory();
    } catch (e) {
      console.error(e);
      alert("Invalid username or password");
    }
  };

  const handleGoogleLogin = async () => {
    if (isLoggingIn) return;
    setIsLoggingIn(true);
    try {
      const result = await signInWithPopup(auth, googleProvider);
      const credential = GoogleAuthProvider.credentialFromResult(result);
      if (result.user) {
        setRecruiter({
          name: result.user.displayName || 'Recruiter',
          email: result.user.email || 'No Email',
          isAdmin: false,
          accessToken: credential?.accessToken || undefined
        });
        setIsAuthed(true);
        fetchHistory();
      }
    } catch (e: any) {
      if (e.code !== 'auth/popup-closed-by-user' && e.code !== 'auth/cancelled-popup-request') {
          console.error("Authentication Error:", e);
          alert("Login failed: " + e.message + " (" + e.code + ")");
      } else {
          console.log("Authentication cancelled by user.");
      }
    } finally {
        setIsLoggingIn(false);
    }
  };

  // Theme helper
  const isDark = theme === 'dark';
  const bg = isDark ? 'bg-slate-950' : 'bg-slate-50';
  const text = isDark ? 'text-slate-100' : 'text-slate-900';
  const cardBg = isDark ? 'bg-slate-900' : 'bg-white';
  const borderColor = isDark ? 'border-slate-800' : 'border-slate-200';
  const textColorMuted = isDark ? 'text-slate-400' : 'text-slate-500';
  const inputBg = isDark ? 'bg-slate-950' : 'bg-white';
  const inputBorder = isDark ? 'border-slate-700' : 'border-slate-200';
  const inputFocus = isDark ? 'focus:ring-orange-500 focus:border-orange-500' : 'focus:ring-indigo-500 focus:border-indigo-300';

  useEffect(() => {
    const init = async () => {
        try {
            await signInAnonymously(auth);
            setIsAuthed(true);
        } catch (e) {
            console.warn("Auth failed or unsupported, proceeding without auth. Firebase features (history) may not work.", e);
        }
    };
    init();
  }, []);

  const evalResults = (results as any)?.candidates ? (results as unknown as EvaluationResults) : null;

  // New expanded navigation
  
  // Sidebar layout for non-landing views
  const Layout = ({ children }: { children: React.ReactNode }) => (
    <div className={`min-h-screen ${bg} ${text} flex font-sans`}>
      <aside className={`w-64 border-r ${borderColor} p-6 flex flex-col gap-8`}>
         <div className="flex flex-col">
            <div className="text-xl font-black">NextGen<span className="text-orange-500">ATS</span></div>
            <div className="text-[10px] font-bold text-orange-300 uppercase tracking-widest">By Mentors Eduserv</div>
         </div>
         <nav className="flex flex-col gap-2">
            {[ {v: 'ats', label: 'Screening'}, {v: 'shortlist', label: 'Shortlist'}, {v: 'pipeline', label: 'Pipeline'}, {v: 'dashboard', label: 'Dashboard'}, {v: 'history', label: 'History'}, {v: 'admin', label: 'Admin'} ].map(item => (
                <button key={item.v} onClick={() => setView(item.v as any)} className={`px-4 py-2 rounded-xl text-sm font-semibold transition-colors ${view === item.v ? 'bg-indigo-600 text-white' : 'hover:bg-slate-800/10'}`}>
                    {item.label}
                </button>
            ))}
         </nav>
      </aside>
      <main className="flex-1 p-8 overflow-y-auto">
        {children}
      </main>
    </div>
  );


  const saveEvaluation = async (evaluation: EvaluationResults) => {
    const userId = recruiter?.email || auth.currentUser?.uid;
    if (!userId) return;
    await addDoc(collection(db, 'evaluations'), {
      userId: userId,
      jobDescription,
      ...evaluation,
      createdAt: new Date().toISOString()
    });
  };

  const handleSaveDraft = async () => {
    const userId = recruiter?.email || auth.currentUser?.uid;
    if (!userId) {
       alert("Please login first to save draft.");
       return;
    }
    
    try {
        await addDoc(collection(db, 'evaluations'), {
          userId: userId,
          jobDescription,
          jobRole,
          status: 'draft',
          createdAt: new Date().toISOString()
        });
        alert("Draft saved!");
    } catch(e) {
        console.error(e);
        alert("Failed to save draft.");
    }
  };

  const fetchHistory = async () => {
    const userId = recruiter?.email || auth.currentUser?.uid;
    if (!userId) return;
    const q = query(collection(db, 'evaluations'), where('userId', '==', userId));
    const querySnapshot = await getDocs(q);
    const historyData = querySnapshot.docs.map(doc => doc.data());
    setHistory(historyData);
  };

  useEffect(() => {
    let url: string | null = null;
    if (selectedCandidate) {
        const resumeFile = resumes.find(r => r.file.name === selectedCandidate.filename)?.file;
        if (resumeFile) {
            url = URL.createObjectURL(resumeFile);
        }
    }
    
    setResumeUrl(prev => {
        if (prev) URL.revokeObjectURL(prev);
        return url;
    });
    
    return () => {
        if (url) URL.revokeObjectURL(url);
    };
  }, [selectedCandidate, resumes]);


  const PDFPreview = ({ url }: { url: string }) => {
      const canvasRef = useRef<HTMLCanvasElement>(null);
      useEffect(() => {
          if (!url || !canvasRef.current) return;
          
          let isCancelled = false;
          
          const renderPDF = async () => {
              try {
                  const loadingTask = pdfjsLib.getDocument(url);
                  const pdf = await loadingTask.promise;
                  const page = await pdf.getPage(1); // Preview first page
                  const viewport = page.getViewport({ scale: 1.0 });
                  
                  const canvas = canvasRef.current;
                  if (!canvas) return;
                  
                  const context = canvas.getContext('2d');
                  if (!context) return;
                  
                  canvas.height = viewport.height;
                  canvas.width = viewport.width;
                  
                  if (isCancelled) return;
                  
                  await page.render({ canvasContext: context, viewport }).promise;
              } catch (e) {
                  console.error("Error rendering PDF:", e);
              }
          };
          
          renderPDF();
          
          return () => { isCancelled = true; };
      }, [url]);
      
      return <canvas ref={canvasRef} className="w-full h-auto rounded-2xl border border-slate-200" />;
  };

  const handleResumeUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setResumes(Array.from(e.target.files).map(file => ({file, status: 'Pending'})));
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const files = Array.from(e.dataTransfer.files).filter((f: File) => f.type === 'application/pdf');
      if (files.length > 0) {
        setResumes(files.map(file => ({file, status: 'Pending'})));
      }
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const processResumesSimple = async () => {
    setLoading(true);
    setResumes(prev => prev.map(r => ({...r, status: 'Parsing'})));
    try {
        const parsedResumes = await Promise.all(resumes.map(async (resume) => {
            const { file } = resume;
             const arrayBuffer = await file.arrayBuffer();
             const pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;
             let text = '';
             for (let i = 1; i <= pdf.numPages; i++) {
                 const page = await pdf.getPage(i);
                 const content = await page.getTextContent();
                 text += content.items.map((item: any) => item.str).join(' ');
             }
             return {filename: file.name, text};
        }));
        setResumes(prev => prev.map(r => ({...r, status: 'Parsed'})));
        setResults({
            candidates: parsedResumes.map(r => ({
                name: r.filename.replace('.pdf', ''),
                filename: r.filename,
                score: 0,
                skills: [],
                education: [],
                workExperience: [],
                portfolioLink: undefined,
                yearsOfExperience: 0,
                justification: 'Quick upload, no AI analysis.',
                verdict: 'Average'
            })),
            insights: 'Quick upload mode.',
            shortlistQuality: 'C',
            matchingCriteria: []
        });
    } catch (e) {
        setResumes(prev => prev.map(r => ({...r, status: 'Error'})));
        console.error(e);
    }
    setLoading(false);
  };

  const processResumesAI = async () => {
    setLoading(true);
    setResumes(prev => prev.map(r => ({...r, status: 'Parsing'})));
    let parsedResumes: {filename: string, text: string}[];
    try {
        parsedResumes = await Promise.all(resumes.map(async (resume) => {
            const { file } = resume;
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;
            let text = '';
            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const content = await page.getTextContent();
                text += content.items.map((item: any) => item.str).join(' ');
            }
            return {filename: file.name, text};
        }));
        setResumes(prev => prev.map(r => ({...r, status: 'Parsed'})));
    } catch (e) {
        setResumes(prev => prev.map(r => ({...r, status: 'Error'})));
        console.error(e);
        setLoading(false);
        return;
    }

    // 2. Call Gemini to rank them
    const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY! });
    
    const prompt = `
      Evaluate the following candidates based on the job description.
      
      Job Description:
      ${jobDescription}
      
      Candidates:
      ${parsedResumes.map(r => `--- Resume: ${r.filename} ---\n${r.text}`).join('\n')}
      
      Tasks:
      1. Extract candidate name, email, phone (if present, otherwise leave off), location (if present, otherwise leave off), portfolio link (if present, otherwise leave off), score (0-100), years of experience (number), key matching skills, list of education (degrees/institutions), list of work experience (companies/roles), a 1-sentence justification, and a verdict.
      2. Provide a 2-sentence summary of overall 'intelligence' insights for this role.
      3. Assign an overall shortlist quality grade (A+, A, B, or C).
      4. Extract 4-5 key matching criteria from the JD.
    `;

    try {
      let response;
      let retries = 0;
      while (retries < 3) {
        try {
          response = await ai.models.generateContent({
            model: "gemini-3.1-pro-preview",
            contents: prompt,
            config: {
              responseMimeType: "application/json",
              responseSchema: {
                type: Type.OBJECT,
                properties: {
                  candidates: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        name: { type: Type.STRING },
                        email: { type: Type.STRING },
                        phone: { type: Type.STRING },
                        location: { type: Type.STRING },
                        portfolioLink: { type: Type.STRING },
                        filename: { type: Type.STRING },
                        score: { type: Type.NUMBER },
                        yearsOfExperience: { type: Type.NUMBER },
                        skills: { type: Type.ARRAY, items: { type: Type.STRING } },
                        education: { type: Type.ARRAY, items: { type: Type.STRING } },
                        workExperience: { type: Type.ARRAY, items: { type: Type.STRING } },
                        justification: { type: Type.STRING },
                        verdict: { type: Type.STRING, enum: ['Excellent', 'Good', 'Average', 'Poor'] }
                      },
                      required: ['name', 'filename', 'score', 'yearsOfExperience', 'skills', 'education', 'workExperience', 'justification', 'verdict', 'location']
                    }
                  },
                  insights: { type: Type.STRING },
                  shortlistQuality: { type: Type.STRING, enum: ['A+', 'A', 'B', 'C'] },
                  matchingCriteria: { type: Type.ARRAY, items: { type: Type.STRING } }
                },
                required: ['candidates', 'insights', 'shortlistQuality', 'matchingCriteria']
              }
            }
          });
          break; // Success
        } catch (error: any) {
            if ((error?.error?.code === 429 || error?.status === 429) && retries < 2) {
                retries++;
                await new Promise(resolve => setTimeout(resolve, Math.pow(2, retries) * 1000));
                continue;
            }
            throw error;
        }
      }
      
      if (!response) throw new Error("Failed to get response from Gemini");

      const parsedData = JSON.parse(response.text);
      setResults(parsedData);
      await saveEvaluation(parsedData);
    } catch (error) {
      console.error("Gemini Error:", error);
      alert("Failed to analyze resumes. The API quota was exceeded. Please try again or check your billing plan.");
    } finally {
      setLoading(false);
    }
  };

  const downloadAsPDF = () => {
    if (!evalResults) return;
    const sortedCandidates = [...evalResults.candidates].sort((a, b) => b.score - a.score);
    const doc = new jsPDF();
    doc.setFontSize(22);
    doc.text("Mentors Eduserv - ATS Evaluation Report", 14, 22);
    doc.setFontSize(14);
    doc.text(`Role: ${jobRole}`, 14, 30);
    doc.setFontSize(11);
    doc.text(`Quality Grade: ${evalResults.shortlistQuality}`, 14, 38);
    const insightsText = doc.splitTextToSize(`Insights: ${evalResults.insights}`, 180);
    doc.text(insightsText, 14, 46);
    const startY = 46 + (insightsText.length * 7) + 10;

    const tableData = sortedCandidates.map(c => [
      c.name,
      c.email || 'N/A',
      c.phone || 'N/A',
      c.location || 'N/A',
      c.portfolioLink || 'N/A',
      c.yearsOfExperience + " yrs",
      c.score + "%",
      c.verdict,
      c.skills.join(", "),
      c.justification
    ]);

    autoTable(doc, {
      startY: startY,
      head: [["Name", "Email", "Phone", "Location", "Portfolio", "Exp", "Score", "Verdict", "Skills", "Justification"]],
      body: tableData,
      styles: { fontSize: 9, overflow: 'linebreak', cellPadding: 1 },
      columnStyles: {
        0: { cellWidth: 25 }, // Name
        1: { cellWidth: 20 }, // Email
        2: { cellWidth: 15 }, // Phone
        3: { cellWidth: 20 }, // Location
        4: { cellWidth: 20 }, // Portfolio
        5: { cellWidth: 10 }, // Exp
        6: { cellWidth: 15 }, // Score
        7: { cellWidth: 15 }, // Verdict
        8: { cellWidth: 25 }, // Skills
        9: { cellWidth: 35 }, // Justification
      },
      headStyles: { fontSize: 9, fillColor: [63, 81, 181] },
      margin: { left: 5, right: 5 },
      didDrawCell: (data) => {
        if (data.section === 'body' && data.column.index === 4 && data.cell.raw !== 'N/A') {
            const url = data.cell.raw as string;
            const fullUrl = url.startsWith('http') ? url : `https://${url}`;
            doc.link(data.cell.x, data.cell.y, data.cell.width, data.cell.height, { url: fullUrl });
        }
      }
    });

    doc.save("Mentors_Eduserv_Evaluation_Report.pdf");
  };

  const downloadAsExcel = () => {
    if (!evalResults) return;
    const sortedCandidates = [...evalResults.candidates].sort((a, b) => b.score - a.score);
    const ws = XLSX.utils.json_to_sheet(sortedCandidates.map(c => ({
      Name: c.name,
      Email: c.email || '',
      Phone: c.phone || '',
      Location: c.location || 'N/A',
      Score: c.score,
      Verdict: c.verdict,
      Portfolio: c.portfolioLink || '',
      Skills: c.skills.join(", "),
      Justification: c.justification,
      Filename: c.filename
    })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Shortlist");
    XLSX.writeFile(wb, "Mentors_Eduserv_Evaluation_Shortlist.xlsx");
  };

  const downloadAsWord = () => {
    if (!evalResults) return;
    const sortedCandidates = [...evalResults.candidates].sort((a, b) => b.score - a.score);

    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: "Mentors Eduserv - ATS Candidate Evaluation Report",
                bold: true,
                size: 32,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "Shortlist Quality: ", bold: true }),
              new TextRun(evalResults.shortlistQuality),
            ],
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "Insights: ", bold: true }),
              new TextRun(evalResults.insights),
            ],
            spacing: { after: 400 },
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: ["Name", "Email", "Phone", "Location", "Portfolio", "Score", "Verdict", "Justification"].map(text => 
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text, bold: true })] })],
                    shading: { fill: "F3F4F6" }
                  })
                ),
              }),
              ...sortedCandidates.map(c => new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph(c.name)] }),
                  new TableCell({ children: [new Paragraph(c.email || 'N/A')] }),
                  new TableCell({ children: [new Paragraph(c.phone || 'N/A')] }),
                  new TableCell({ children: [new Paragraph(c.location || 'N/A')] }),
                  new TableCell({ children: [new Paragraph(c.portfolioLink || 'N/A')] }),
                  new TableCell({ children: [new Paragraph(c.score.toString() + "%")] }),
                  new TableCell({ children: [new Paragraph(c.verdict)] }),
                  new TableCell({ children: [new Paragraph(c.justification)] }),
                ],
              })),
            ],
          }),
        ],
      }],
    });

    Packer.toBlob(doc).then(blob => {
      saveAs(blob, "Mentors_Eduserv_Evaluation_Report.docx");
    });
  };

  if (view === 'landing') {
    return (
      <div className={`min-h-screen ${isDark ? 'bg-slate-950 text-white' : 'bg-slate-50 text-slate-900'} relative overflow-hidden`}>
        {/* Floating Background Effects */}
        {[...Array(6)].map((_, i) => (
          <motion.div
            key={i}
            className={`absolute rounded-full ${isDark ? 'bg-orange-500 opacity-10' : 'bg-indigo-300 opacity-20'} blur-3xl`}
            animate={{
              x: [Math.random() * 1200 - 600, Math.random() * 1200 - 600],
              y: [Math.random() * 1000 - 500, Math.random() * 1000 - 500],
              scale: [1, 1.5, 1],
            }}
            transition={{ duration: 15 + i * 2, repeat: Infinity, ease: "linear" }}
            style={{ width: `${300 + i * 100}px`, height: `${300 + i * 100}px` }}
          />
        ))}

        <nav className="absolute top-0 w-full p-8 flex justify-between items-center z-20">
          <div className="flex flex-col">
          <div className="text-2xl font-black">NextGen<span className="text-orange-500">ATS</span></div>
          <div className="text-[10px] font-bold text-orange-300 uppercase tracking-widest">By Mentors Eduserv</div>
         </div>
          <div className="flex gap-4 items-center">
            {recruiter ? (
                <div className="flex items-center gap-4">
                    <div className="flex items-center gap-2">
                        <User className="w-6 h-6 text-orange-500" />
                        <div className="flex flex-col text-right">
                            <span className="text-sm font-bold">{recruiter.name} {recruiter.isAdmin && '(Admin)'}</span>
                            <span className="text-[10px] text-slate-500">{recruiter.email}</span>
                        </div>
                    </div>
                    <button onClick={() => setView('ats')} className={`px-5 py-2 rounded-full bg-orange-500 text-white hover:bg-orange-600 transition`}>Enter ATS</button>
                    <button onClick={handleLogout} className="p-2 rounded-full hover:bg-slate-800 transition">
                         <LogOut className="w-5 h-5 text-slate-400 hover:text-red-500" />
                    </button>
                </div>
            ) : (
                <div className="flex gap-2">
                    {showLoginSelection ? (
                        <>
                            <button onClick={() => {setLoginMode('google');}} className={`px-5 py-2 rounded-full ${loginMode === 'google' ? 'bg-orange-500 text-white' : 'border border-orange-500 text-orange-500'} transition`}>Google</button>
                            <button onClick={() => {setLoginMode('recruiter');}} className={`px-5 py-2 rounded-full ${loginMode === 'recruiter' ? 'bg-orange-500 text-white' : 'border border-orange-500 text-orange-500'} transition`}>Recruiter</button>
                        </>
                    ) : (
                        <button onClick={() => setShowLoginSelection(true)} className={`px-5 py-2 rounded-full bg-orange-500 text-white transition`}>Login</button>
                    )}
                </div>
            )}
            <button onClick={() => setView('history')} className={`px-5 py-2 rounded-full border border-orange-500 text-orange-500 hover:bg-orange-600 hover:text-white transition`}>History</button>
            <button onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')} className={`px-5 py-2 rounded-full border border-orange-500 text-orange-500 hover:bg-orange-600 hover:text-white transition`}>
              {theme === 'dark' ? 'Light' : 'Dark'} Mode
            </button>
          </div>
        </nav>

        <main className="relative z-10">
          {/* Hero */}
          <section className="min-h-screen flex flex-col items-center justify-center text-center px-4 space-y-8">
            <motion.h1 
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="text-7xl md:text-9xl font-black tracking-tighter"
            >
              Hire Smarter, <br/><span className="text-orange-500">Not Harder.</span>
            </motion.h1>
            <p className={`text-xl ${isDark ? 'text-slate-400' : 'text-slate-600'} max-w-2xl`}>
              AI-powered resume screening, ranking, and insights designed to help recruiting teams scale faster. Stop reading hundreds of resumes and start interviewing the best candidates.
            </p>
            <button 
              onClick={
                  recruiter ? () => setView('ats') : 
                  (showLoginSelection && loginMode ? (loginMode === 'google' ? handleGoogleLogin : handleCustomLogin) : () => setShowLoginSelection(true))
              }
              className="px-8 py-4 bg-orange-600 rounded-full font-bold text-lg hover:bg-orange-700 transition"
            >
              {recruiter ? 'Enter ATS' : (showLoginSelection && loginMode ? (loginMode === 'google' ? 'Google Login' : 'Login') : 'Login')}                
            </button>
            {showLoginSelection && !loginMode && (
                <div className="flex gap-2 justify-center mt-4">
                    <button onClick={() => setLoginMode('google')} className="px-5 py-2 rounded-full border border-orange-500 text-orange-500">Google</button>
                    <button onClick={() => setLoginMode('recruiter')} className="px-5 py-2 rounded-full border border-orange-500 text-orange-500">Recruiter</button>
                </div>
            )}
            {loginMode === 'recruiter' && !recruiter && (
                <div className="flex flex-col gap-2 mt-4 max-w-sm mx-auto">
                    <input type="text" placeholder="Username" onChange={e => setLoginUsername(e.target.value)} className="px-4 py-2 rounded-full border border-slate-700 bg-slate-900"/>
                    <input type="password" placeholder="Password" onChange={e => setLoginPassword(e.target.value)} className="px-4 py-2 rounded-full border border-slate-700 bg-slate-900"/>
                </div>
            )}
          </section>

          {/* Features */}
          <section className={`py-24 px-8 ${isDark ? 'bg-slate-900/50' : 'bg-white/50'}`}>
            <div className="max-w-6xl mx-auto grid md:grid-cols-3 gap-12">
              {[
                { title: "AI-Powered Ranking", desc: "Instantly rank candidates against job descriptions." },
                { title: "Skill Extraction", desc: "Automated analysis of candidate competencies." },
                { title: "Professional Reporting", desc: "Download ready-to-use reports for stakeholders." }
              ].map((f, i) => (
                <div key={i} className={`p-8 ${isDark ? 'bg-slate-950 border-slate-800' : 'bg-white border-slate-200'} rounded-3xl border space-y-4`}>
                  <h3 className="text-xl font-bold">{f.title}</h3>
                  <p className={`${textColorMuted} text-sm`}>{f.desc}</p>
                </div>
              ))}
            </div>
          </section>

          {/* Professional Showcase */}
          <section className="py-24 px-8">
            <div className="max-w-6xl mx-auto">
              <h2 className="text-4xl font-bold text-center mb-16">Trusted by Industry Leaders</h2>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-8 opacity-60">
                  {['NextGen Edu', 'Mentors Labs', 'CareerFlow', 'SkillMatch'].map(brand => (
                    <div key={brand} className="text-center font-bold text-lg">{brand}</div>
                 ))}
              </div>
            </div>
          </section>

          {/* CTA Footer */}
          <section className="py-24 text-center space-y-6">
            <h2 className="text-4xl font-bold">Ready to modernize your recruiting?</h2>
            <button 
              onClick={recruiter ? () => setView('ats') : handleGoogleLogin}
              className={`px-8 py-4 ${isDark ? 'bg-white text-slate-950 hover:bg-slate-200' : 'bg-slate-950 text-white hover:bg-slate-800'} rounded-full font-bold text-lg transition`}
            >
              Launch ATS Engine
            </button>
          </section>
        </main>
      </div>
    );
  }

  if (view === 'history') {
    const totalEvaluations = history.length;
    const totalResumes = history.reduce((acc, h) => acc + (h.candidates?.length || 0), 0);

    return (
      <div className={`min-h-screen ${bg} p-8 font-sans ${text}`}>
        <header className={`mb-8 border-b ${borderColor} pb-6 flex items-center justify-between`}>
          <div className="flex items-center gap-4">
            <button onClick={() => setView('landing')} className={`${textColorMuted} hover:text-orange-500 transition-colors`}>&larr; Back to Landing</button>
            <h1 className="text-3xl font-bold">Activity Dashboard</h1>
          </div>
          {recruiter && (
             <div className="flex items-center gap-3">
                <div className="flex items-center gap-2">
                    <User className="w-6 h-6 text-orange-500" />
                    <div className="flex flex-col text-right">
                        <span className="text-sm font-bold">{recruiter.name}</span>
                        <span className="text-[10px] text-slate-500">{recruiter.email}</span>
                    </div>
                </div>
                <button onClick={handleLogout} className={`p-2.5 ${cardBg} border ${borderColor} rounded-xl`}>
                    <LogOut className="w-5 h-5 text-slate-400 hover:text-red-500" />
                </button>
             </div>
          )}
        </header>

        <div className="grid grid-cols-2 gap-4 mb-8">
            <div className={`${cardBg} border ${borderColor} p-6 rounded-2xl`}>
                <div className={`${textColorMuted} text-xs font-bold uppercase`}>Total Evaluations</div>
                <div className="text-4xl font-black mt-2">{totalEvaluations}</div>
            </div>
            <div className={`${cardBg} border ${borderColor} p-6 rounded-2xl`}>
                <div className={`${textColorMuted} text-xs font-bold uppercase`}>Candidates Analyzed</div>
                <div className="text-4xl font-black mt-2">{totalResumes}</div>
            </div>
        </div>

        <div className="grid gap-4">
          {history.length === 0 && <p className={textColorMuted}>No evaluations found.</p>}
          {history.map((h, i) => (
             <div key={i} className={`${cardBg} border ${borderColor} p-6 rounded-2xl`}>
               <h3 className="font-bold text-lg mb-2">{h.createdAt ? new Date(h.createdAt).toLocaleDateString() : 'Unknown date'}</h3>
               <p className={`text-sm ${textColorMuted} line-clamp-2`}>{h.jobDescription}</p>
               <p className="text-xs text-orange-400 mt-2">{h.candidates.length} candidates analyzed</p>
             </div>
          ))}
        </div>
      </div>
    );
  }

  return (
    <div className={`min-h-screen ${bg} p-8 font-sans ${text}`}>
      <header className={`mb-8 border-b ${borderColor} pb-6 flex items-center justify-between`}>
        <div className='flex items-center gap-4'>
            <button onClick={() => setView('landing')} className={`${textColorMuted} hover:text-indigo-600`}>&larr; Back</button>
            <div>
              <h1 className="text-3xl font-bold tracking-tight">ATS Resume Screener</h1>
              <p className={`text-sm ${textColorMuted} mt-1`}>Intelligent candidate matching engine</p>
            </div>
        </div>
        <div className="flex gap-3">
          <div className={`flex ${cardBg} border ${borderColor} rounded-xl overflow-hidden shadow-sm`}>
            <button
               onClick={() => setMode('ai')}
               className={`px-4 py-2 text-xs font-bold ${mode === 'ai' ? 'bg-indigo-600 text-white' : 'hover:bg-slate-800/10'} border-r border-slate-800/10 flex items-center gap-2`}
            >
              AI Ranked
            </button>
            <button
               onClick={() => setMode('simple')}
               className={`px-4 py-2 text-xs font-bold ${mode === 'simple' ? 'bg-indigo-600 text-white' : 'hover:bg-slate-800/10'} border-r border-slate-800/10 flex items-center gap-2`}
            >
              Quick Upload
            </button>
            <button
               onClick={() => setMode('import')}
               className={`px-4 py-2 text-xs font-bold ${mode === 'import' ? 'bg-indigo-600 text-white' : 'hover:bg-slate-800/10'} flex items-center gap-2`}
            >
              Import Data
            </button>
          </div>
          <button onClick={() => setView('history')} className={`px-5 py-2.5 ${cardBg} border ${borderColor} rounded-xl text-sm font-semibold flex items-center gap-2`}>
            <History className="w-4 h-4" /> History
          </button>
          {recruiter && (
             <div className="flex items-center gap-3">
                <div className="flex items-center gap-2">
                    <User className="w-6 h-6 text-orange-500" />
                    <div className="flex flex-col text-right">
                        <span className="text-sm font-bold">{recruiter.name}</span>
                        <span className="text-[10px] text-slate-500">{recruiter.email}</span>
                    </div>
                </div>
                <button onClick={handleLogout} className={`p-2.5 ${cardBg} border ${borderColor} rounded-xl`}>
                    <LogOut className="w-5 h-5 text-slate-400 hover:text-red-500" />
                </button>
             </div>
          )}
          
          {evalResults && evalResults.candidates && (
            <div className={`flex ${cardBg} border ${borderColor} rounded-xl overflow-hidden shadow-sm`}>
              {selectedCandidateFilenames.length > 0 ? (
                  <button type="button" onClick={() => {
                        const selected = evalResults.candidates.filter(c => selectedCandidateFilenames.includes(c.filename));
                        // Re-use downloadAsExcel logic, just filter by selected first
                        const ws = XLSX.utils.json_to_sheet(selected.map(c => ({
                            Name: c.name,
                            Email: c.email || '',
                            Phone: c.phone || '',
                            Score: c.score,
                            Verdict: c.verdict,
                            Skills: c.skills.join(", "),
                            Justification: c.justification,
                            Filename: c.filename
                        })));
                        const wb = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(wb, ws, "Shortlist");
                        XLSX.writeFile(wb, "Selected_Candidates_Export.xlsx");
                  }} className="px-4 py-2 text-xs font-bold bg-indigo-100 text-indigo-700 hover:bg-indigo-200 border-r border-slate-800/10 flex items-center gap-2">
                    <Download className="w-3 h-3" /> Export Selected ({selectedCandidateFilenames.length})
                  </button>
              ) : null}
              <button type="button" onClick={downloadAsPDF} className="px-4 py-2 text-xs font-bold hover:bg-slate-800/10 border-r border-slate-800/10 flex items-center gap-2">
                <Download className="w-3 h-3" /> PDF
              </button>
              <button type="button" onClick={downloadAsExcel} className="px-4 py-2 text-xs font-bold hover:bg-slate-800/10 border-r border-slate-800/10 flex items-center gap-2">
                <Download className="w-3 h-3" /> EXCEL
              </button>
              <button type="button" onClick={downloadAsWord} className="px-4 py-2 text-xs font-bold hover:bg-slate-800/10 flex items-center gap-2">
                <Download className="w-3 h-3" /> WORD
              </button>
            </div>
          )}
          <button type="button" onClick={handleSaveDraft} className={`px-5 py-2.5 ${cardBg} border ${borderColor} rounded-xl text-sm font-semibold transition-colors`}>Save Draft</button>
          <button 
             type="button"
             onClick={mode === 'ai' ? processResumesAI : processResumesSimple}
             disabled={loading || (mode === 'ai' && !jobDescription) || resumes.length === 0}
             className="px-5 py-2.5 bg-indigo-600 text-white rounded-xl text-sm font-semibold shadow-lg hover:bg-indigo-700 transition-colors disabled:opacity-50"
          >
            {loading ? 'Processing...' : (mode === 'ai' ? 'Run Analysis' : 'Upload & View')}
          </button>
        </div>
      </header>

        {mode === 'ai' && (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
             <div className={`${cardBg} p-6 rounded-2xl shadow-sm border ${borderColor}`}>
             <label className="block text-sm font-bold mb-3">Job Role</label>
             <input
                className={`w-full p-4 border ${inputBorder} rounded-xl ${inputBg} ${isDark ? 'text-white' : ''} ${inputFocus} transition-all mb-4`}
                value={jobRole}
                onChange={(e) => setJobRole(e.target.value)}
                placeholder="e.g. Senior Graphic Designer"
            />
            <label className="block text-sm font-bold mb-3">Job Description</label>
            <textarea
                className={`w-full h-80 p-4 border ${inputBorder} rounded-xl ${inputBg} ${isDark ? 'text-white' : ''} ${inputFocus} transition-all`}
                value={jobDescription}
                onChange={(e) => setJobDescription(e.target.value)}
                placeholder="Paste the job description here..."
            />
            </div>
            <div className={`${cardBg} p-6 rounded-2xl shadow-sm border ${borderColor}`}>
            <label className="block text-sm font-bold mb-3">Upload Resumes</label>
            <div 
                onDragOver={handleDragOver}
                onDrop={handleDrop}
                className={`relative border-2 border-dashed ${inputBorder} rounded-2xl p-8 flex flex-col items-center justify-center ${textColorMuted} hover:border-orange-500 hover:bg-slate-950/50 transition-all cursor-pointer`}
            >
                <input 
                type="file" 
                multiple 
                accept=".pdf" 
                onChange={handleResumeUpload} 
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" 
                id="resume-upload" 
                />
                <FileText className="w-8 h-8 mb-4" />
                <span className="font-semibold text-sm">Click to upload or drag and drop</span>
                <p className="text-xs mt-2">PDF files only ({resumes.length} selected)</p>
                {resumes.length > 0 && (
                    <div className="w-full mt-4 space-y-2">
                         {resumes.map((r, i) => (
                             <div key={i} className="flex justify-between items-center bg-slate-800/10 p-2 rounded text-xs">
                                 <span className="truncate max-w-[150px]">{r.file.name}</span>
                                 <span className={`px-2 py-0.5 rounded ${r.status === 'Pending' ? 'bg-slate-200' : r.status === 'Parsing' ? 'bg-blue-200' : r.status === 'Parsed' ? 'bg-green-200' : 'bg-red-200'}`}> {r.status} </span>
                             </div>
                         ))}
                    </div>
                )}
            </div>
            </div>
        </div>
        )}
        
        {mode === 'simple' && (
        <div className="grid grid-cols-1 gap-6 mb-8">
            <div className={`${cardBg} p-6 rounded-2xl shadow-sm border ${borderColor}`}>
                <label className="block text-sm font-bold mb-3">Upload Resumes (Simple)</label>
                <div 
                    onDragOver={handleDragOver}
                    onDrop={handleDrop}
                    className={`relative border-2 border-dashed ${inputBorder} rounded-2xl p-8 flex flex-col items-center justify-center ${textColorMuted} hover:border-orange-500 hover:bg-slate-950/50 transition-all cursor-pointer`}
                >
                    <input 
                    type="file" 
                    multiple 
                    accept=".pdf" 
                    onChange={handleResumeUpload} 
                    className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" 
                    id="resume-upload" 
                    />
                    <FileText className="w-8 h-8 mb-4" />
                    <span className="font-semibold text-sm">Click to upload or drag and drop</span>
                    <p className="text-xs mt-2">PDF files only ({resumes.length} selected)</p>
                    {resumes.length > 0 && (
                        <div className="w-full mt-4 space-y-2">
                             {resumes.map((r, i) => (
                                 <div key={i} className="flex justify-between items-center bg-slate-800/10 p-2 rounded text-xs">
                                     <span className="truncate max-w-[150px]">{r.file.name}</span>
                                     <span className={`px-2 py-0.5 rounded ${r.status === 'Pending' ? 'bg-slate-200' : r.status === 'Parsing' ? 'bg-blue-200' : r.status === 'Parsed' ? 'bg-green-200' : 'bg-red-200'}`}> {r.status} </span>
                                 </div>
                             ))}
                        </div>
                    )}
                </div>
            </div>
        </div>
        )}
        
        {mode === 'import' && (
        <div className="grid grid-cols-1 gap-6 mb-8">
            <div className={`${cardBg} p-6 rounded-2xl shadow-sm border ${borderColor}`}>
                <label className="block text-sm font-bold mb-3">Upload Candidates Excel File</label>
                <div 
                    className={`relative border-2 border-dashed ${inputBorder} rounded-2xl p-8 flex flex-col items-center justify-center ${textColorMuted} hover:border-orange-500 hover:bg-slate-950/50 transition-all cursor-pointer`}
                >
                    <input 
                    type="file" 
                    accept=".xlsx, .xls" 
                    onChange={(e) => {
                       const file = e.target.files?.[0];
                       if (!file) return;

                       const reader = new FileReader();
                       reader.onload = (evt) => {
                           const data = evt.target?.result;
                           const workbook = XLSX.read(data, {type: 'binary'});
                           const sheetName = workbook.SheetNames[0];
                           const worksheet = workbook.Sheets[sheetName];
                           const json = XLSX.utils.sheet_to_json(worksheet);

                           const candidates = json.map((row: any) => ({
                               name: row.Name || 'Unknown',
                               email: row.Email,
                               phone: row.Phone,
                               jobProfile: row.JobProfile || 'N/A', // Capturing Job Profile
                               score: Number(row.Score) || 0,
                               verdict: row.Verdict || 'Average',
                               skills: String(row.Skills || '').split(',').map((s: string) => s.trim()).filter(Boolean),
                               justification: row.Justification || 'Imported from Excel',
                               resumeUrl: row.ResumeUrl || null // Capturing actual Resume URL
                           }));

                           setResults({
                               candidates,
                               insights: 'Imported from Excel',
                               shortlistQuality: 'B',
                               matchingCriteria: []
                           });
                       };
                       reader.readAsBinaryString(file);
                    }} 
                    className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" 
                    id="excel-upload" 
                    />
                    <FileText className="w-8 h-8 mb-4" />
                    <span className="font-semibold text-sm">Click to upload Excel file</span>
                </div>
            </div>
        </div>
        )}
      
      {results && results.candidates && (
        <div className="space-y-6">
          {/* Matching Criteria Bar */}
          <section className={`${isDark ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'} border-y px-8 py-4 flex items-center gap-6 -mx-8 overflow-x-auto`}>
            <div className="flex items-center gap-2 shrink-0">
              <span className={`text-[10px] font-bold ${textColorMuted} uppercase tracking-widest flex items-center gap-1.5`}>
                <div className="w-1.5 h-1.5 bg-indigo-500 rounded-full animate-pulse" />
                Matching Engine
              </span>
              <div className={`w-px h-4 ${borderColor}`} />
            </div>
            <div className="flex gap-2">
              {evalResults.matchingCriteria.map((tag: string, idx: number) => (
                <span key={idx} className={`px-3 py-1 ${isDark ? 'bg-indigo-950 border-indigo-800 text-indigo-300' : 'bg-indigo-50 border-indigo-100 text-indigo-700'} text-[10px] font-bold rounded-full whitespace-nowrap`}>
                  {tag}
                </span>
              ))}
            </div>
          </section>

          <div className="grid grid-cols-1 md:grid-cols-12 gap-6">
            {/* Main Table Area */}
            <div className={`md:col-span-8 ${cardBg} rounded-2xl shadow-sm border ${borderColor} overflow-hidden`}>
              <div className={`p-6 border-b ${borderColor} flex flex-wrap justify-between items-center ${isDark ? 'bg-slate-950/30' : 'bg-slate-50/50'} gap-4`}>
                <h3 className={`font-bold ${text} text-sm`}>AI Shortlisted Candidates <span className={`ml-2 ${textColorMuted} font-normal`}>({evalResults.candidates.length} total)</span></h3>
                <div className="flex flex-wrap gap-2">
                    <input
                        type="text"
                        placeholder="Skill..."
                        disabled={mode === 'simple'}
                        className={`text-[10px] border ${inputBorder} rounded-lg p-1.5 ${inputBg} ${text} ${inputFocus} outline-none w-24 ${mode === 'simple' ? 'opacity-50 cursor-not-allowed' : ''}`}
                        value={filterQuery}
                        onChange={(e) => setFilterQuery(e.target.value)}
                    />
                    <select
                        disabled={mode === 'simple'}
                        className={`text-[10px] border ${inputBorder} rounded-lg p-1.5 ${inputBg} ${text} ${inputFocus} outline-none ${mode === 'simple' ? 'opacity-50 cursor-not-allowed' : ''}`}
                        value={filterVerdict}
                        onChange={(e) => setFilterVerdict(e.target.value)}
                    >
                        <option value="All">All Verdicts</option>
                        <option value="Excellent">Excellent</option>
                        <option value="Good">Good</option>
                        <option value="Average">Average</option>
                        <option value="Poor">Poor</option>
                    </select>
                    <input
                        type="number"
                        placeholder="Min Score"
                        disabled={mode === 'simple'}
                        className={`text-[10px] border ${inputBorder} rounded-lg p-1.5 ${inputBg} ${text} ${inputFocus} outline-none w-20 ${mode === 'simple' ? 'opacity-50 cursor-not-allowed' : ''}`}
                        value={minScore}
                        onChange={(e) => setMinScore(Number(e.target.value))}
                    />
                </div>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left">
                  <thead>
                    <tr className={`${isDark ? 'bg-slate-950/50' : 'bg-slate-50'} text-[10px] font-bold ${textColorMuted} uppercase tracking-widest`}>
                      <th className="px-6 py-4">
                        <input type="checkbox" checked={selectedCandidateFilenames.length === evalResults.candidates.length} onChange={(e) => {
                            if (e.target.checked) setSelectedCandidateFilenames(evalResults.candidates.map(c => c.filename));
                            else setSelectedCandidateFilenames([]);
                        }}/>
                      </th>
                      <th className="px-6 py-4">Candidate</th>
                      <th className="px-6 py-4">Experience & Fit</th>
                      <th className="px-6 py-4">Location</th>
                      <th className="px-6 py-4 text-center">Score</th>
                      <th className="px-6 py-4 text-right">Actions</th>
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${borderColor}`}>
                    {[...evalResults.candidates]
                        .filter(c => 
                            (mode === 'simple' || filterQuery === '' || c.skills.some(s => s.toLowerCase().includes(filterQuery.toLowerCase()))) &&
                            (mode === 'simple' || filterVerdict === 'All' || c.verdict === filterVerdict) &&
                            (mode === 'simple' || c.score >= minScore)
                        )
                        .sort((a,b) => b.score - a.score).map((candidate, idx) => (
                      <tr key={idx} className={`group hover:${isDark ? 'bg-slate-800/50' : 'bg-slate-50'} transition-colors`}>
                        <td className="px-6 py-5">
                            <input type="checkbox" checked={selectedCandidateFilenames.includes(candidate.filename)} onChange={(e) => {
                                if (e.target.checked) setSelectedCandidateFilenames([...selectedCandidateFilenames, candidate.filename]);
                                else setSelectedCandidateFilenames(selectedCandidateFilenames.filter(f => f !== candidate.filename));
                            }}/>
                        </td>
                        <td className="px-6 py-5">
                          <div className="flex items-center gap-3">
                            <div className={`w-9 h-9 rounded-full ${isDark ? 'bg-slate-800' : 'bg-slate-100'} flex items-center justify-center ${isDark ? 'text-slate-200' : 'text-slate-700'} font-bold text-xs border ${borderColor}`}>
                              {candidate.name.charAt(0)}
                            </div>
                            <div>
                              <p className={`text-sm font-bold ${text} leading-none`}>{candidate.name}</p>
                              <p className={`text-[10px] ${textColorMuted} mt-1`}>Role: {candidate.jobProfile}</p>
                              {(candidate.email || candidate.phone) && (
                                <div className={`text-[10px] ${textColorMuted} mt-1 space-y-0.5`}>
                                    {candidate.email && <p>📧 {candidate.email}</p>}
                                    {candidate.phone && <p>📞 {candidate.phone}</p>}
                                </div>
                              )}
                              {candidate.location && (
                                <p className={`text-[10px] mt-1 ${
                                    (candidate.location.toLowerCase().includes('patna') || candidate.location.toLowerCase().includes('bihar')) 
                                    ? 'text-emerald-500 font-bold' 
                                    : textColorMuted
                                }`}>
                                    📍 {candidate.location}
                                </p>
                              )}
                            </div>
                          </div>
                        </td>
                        <td className="px-6 py-5">
                          <p className={`text-[11px] ${isDark ? 'text-slate-300' : 'text-slate-600'} leading-normal`}>Verdict: <span className="font-bold">{candidate.verdict}</span></p>
                          {candidate.resumeUrl ? (
                            <a href={candidate.resumeUrl} target="_blank" rel="noopener noreferrer" className="text-[10px] text-blue-500 hover:underline flex items-center gap-1 mt-1">
                                <Download className="w-3 h-3"/> Download Resume
                            </a>
                           ) : (
                               <p className={`text-[10px] ${textColorMuted} mt-1`}>No resume link</p>
                           )}
                        </td>
                        <td className="px-6 py-5 text-center">
                          <span className="text-sm font-bold text-orange-600 font-mono tracking-tighter">{candidate.score}%</span>
                        </td>
                        <td className="px-6 py-5 text-right">
                          <button 
                            onClick={() => {
                                if (shortlistedCandidateFilenames.includes(candidate.filename)) {
                                    setShortlistedCandidateFilenames(shortlistedCandidateFilenames.filter(f => f !== candidate.filename));
                                } else {
                                    setShortlistedCandidateFilenames([...shortlistedCandidateFilenames, candidate.filename]);
                                }
                            }}
                            className={`p-2 rounded-full ${shortlistedCandidateFilenames.includes(candidate.filename) ? 'text-yellow-500' : 'text-slate-400'}`}
                          >
                             <Star className="w-5 h-5" fill={shortlistedCandidateFilenames.includes(candidate.filename) ? "currentColor" : "none"}/>
                          </button>
                          <button 
                            onClick={() => setSelectedCandidate(candidate)}
                            className={`p-2 ${textColorMuted} hover:text-orange-500 hover:${isDark ? 'bg-slate-800' : 'bg-slate-100'} rounded-lg transition-all opacity-0 group-hover:opacity-100`}
                          >
                            <Eye className="w-4 h-4" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Sidebar Details Area */}
            <div className="md:col-span-4 space-y-6">
              <div className={`${cardBg} rounded-2xl shadow-sm border ${borderColor} p-6`}>
                <h4 className={`font-bold ${text} text-xs uppercase tracking-widest mb-4 flex items-center gap-2`}>
                  <AlertCircle className="w-4 h-4 text-orange-500" />
                  Matching Intelligence
                </h4>
                <div className="space-y-4">
                  <div className={`p-4 ${inputBg} rounded-xl border ${borderColor}`}>
                    <p className={`text-[11px] ${isDark ? 'text-slate-300' : 'text-slate-600'} leading-relaxed font-medium`}>
                      {evalResults.insights}
                    </p>
                  </div>
                  <div className="space-y-3">
                    <div className="flex justify-between items-end">
                      <span className={`text-[10px] font-bold ${textColorMuted} uppercase tracking-wider`}>JD Alignment</span>
                      <span className={`text-xs font-bold ${text}`}>{Math.round(evalResults.candidates.reduce((a: any,b: any) => a+b.score, 0) / evalResults.candidates.length)}%</span>
                    </div>
                    <div className="w-full bg-slate-800 h-1.5 rounded-full overflow-hidden">
                      <div className="bg-orange-500 h-full rounded-full transition-all duration-1000" style={{ width: `${Math.round(evalResults.candidates.reduce((a: any,b: any) => a+b.score, 0) / evalResults.candidates.length)}%` }} />
                    </div>
                  </div>
                </div>
              </div>

              <div className="bg-indigo-600 rounded-2xl p-6 text-white shadow-xl shadow-indigo-100/50 relative overflow-hidden group">
                <div className="absolute top-0 right-0 p-4 opacity-10 group-hover:scale-110 transition-transform">
                  <Award className="w-16 h-16" />
                </div>
                <h4 className="font-bold text-xs uppercase tracking-widest mb-2 opacity-80">Shortlist Quality</h4>
                <p className="text-[11px] opacity-90 mb-4 leading-relaxed">The AI has analyzed the candidate pool against global standards for this seniority level.</p>
                <div className="flex items-center gap-4">
                  <div className="text-4xl font-bold tracking-tighter">{evalResults.shortlistQuality}</div>
                  <div className="h-10 w-[1px] bg-white/20" />
                  <div className="text-[9px] font-bold leading-tight opacity-70 tracking-widest uppercase">Based on<br />Global Market<br />Standards</div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Shortlist View */}
      {view === 'shortlist' && (
          <Layout>
              <h2 className="text-2xl font-bold mb-6 flex justify-between items-center">
                <span>Candidates Overview</span>
                <div className="flex gap-2">
                    <input 
                      type="text" 
                      placeholder="Spreadsheet ID" 
                      value={spreadsheetIdInput} 
                      onChange={e => setSpreadsheetIdInput(e.target.value)}
                      className="px-4 py-2 rounded-xl text-sm border border-slate-200"
                    />
                    <button onClick={handleExportShortlist} className="text-xs px-4 py-2 rounded-xl bg-green-600 text-white hover:bg-green-700 transition-colors">
                        Export Shortlist to Sheets
                    </button>
                    <button onClick={() => setShowOnlyShortlisted(!showOnlyShortlisted)} className={`text-xs px-4 py-2 rounded-xl transition-colors ${showOnlyShortlisted ? 'bg-yellow-100 text-yellow-700' : 'bg-slate-100 text-slate-700'}`}>
                        {showOnlyShortlisted ? 'Showing Shortlisted Only' : 'Showing All Candidates'}
                    </button>
                </div>
              </h2>
              <div className={`${cardBg} rounded-2xl shadow-sm border ${borderColor} overflow-hidden`}>
                <div className="overflow-x-auto">
                <table className="w-full text-left">
                  <thead>
                    <tr className={`${isDark ? 'bg-slate-950/50' : 'bg-slate-50'} text-[10px] font-bold ${textColorMuted} uppercase tracking-widest`}>
                      <th className="px-6 py-4">Candidate</th>
                      <th className="px-6 py-4">Experience & Fit</th>
                      <th className="px-6 py-4 text-center">Score</th>
                      <th className="px-6 py-4 text-right">Actions</th>
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${borderColor}`}>
                    {evalResults?.candidates
                        .filter(c => !showOnlyShortlisted || shortlistedCandidateFilenames.includes(c.filename))
                        .map((candidate, idx) => (
                      <tr key={idx} className={`group hover:${isDark ? 'bg-slate-800/50' : 'bg-slate-50'} transition-colors`}>
                        <td className="px-6 py-5">
                          <div className="flex items-center gap-3">
                            <div className={`w-9 h-9 rounded-full ${isDark ? 'bg-slate-800' : 'bg-slate-100'} flex items-center justify-center ${isDark ? 'text-slate-200' : 'text-slate-700'} font-bold text-xs border ${borderColor}`}>
                              {candidate.name.charAt(0)}
                            </div>
                            <div>
                              <p className={`text-sm font-bold ${text} leading-none`}>{candidate.name}</p>
                              <p className={`text-[10px] ${textColorMuted} mt-1`}>Role: {candidate.jobProfile}</p>
                              {(candidate.email || candidate.phone) && (
                                <div className={`text-[10px] ${textColorMuted} mt-1 space-y-0.5`}>
                                    {candidate.email && <p>📧 {candidate.email}</p>}
                                    {candidate.phone && <p>📞 {candidate.phone}</p>}
                                </div>
                              )}
                              {candidate.location && (
                                <p className={`text-[10px] mt-1 ${
                                    (candidate.location.toLowerCase().includes('patna') || candidate.location.toLowerCase().includes('bihar')) 
                                    ? 'text-emerald-500 font-bold' 
                                    : textColorMuted
                                }`}>
                                    📍 {candidate.location}
                                </p>
                              )}
                            </div>
                          </div>
                        </td>
                        <td className="px-6 py-5">
                          <p className={`text-[11px] ${isDark ? 'text-slate-300' : 'text-slate-600'} leading-normal`}>Verdict: <span className="font-bold">{candidate.verdict}</span></p>
                          {candidate.resumeUrl ? (
                            <a href={candidate.resumeUrl} target="_blank" rel="noopener noreferrer" className="text-[10px] text-blue-500 hover:underline flex items-center gap-1 mt-1">
                                <Download className="w-3 h-3"/> Download Resume
                            </a>
                           ) : (
                               <p className={`text-[10px] ${textColorMuted} mt-1`}>No resume link</p>
                           )}
                        </td>
                        <td className="px-6 py-5 text-center">
                          <span className="text-sm font-bold text-orange-600 font-mono tracking-tighter">{candidate.score}%</span>
                        </td>
                        <td className="px-6 py-5 text-right flex items-center justify-end gap-2">
                          <button 
                            onClick={() => {
                                if (shortlistedCandidateFilenames.includes(candidate.filename)) {
                                    setShortlistedCandidateFilenames(shortlistedCandidateFilenames.filter(f => f !== candidate.filename));
                                } else {
                                    setShortlistedCandidateFilenames([...shortlistedCandidateFilenames, candidate.filename]);
                                }
                            }}
                            className={`p-2 rounded-full ${shortlistedCandidateFilenames.includes(candidate.filename) ? 'text-yellow-500' : 'text-slate-400'}`}
                          >
                             <Star className="w-5 h-5" fill={shortlistedCandidateFilenames.includes(candidate.filename) ? "currentColor" : "none"}/>
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
              </div>
          </Layout>
      )}

      {/* Pipeline View */}
      {view === 'pipeline' && (
          <Layout>
              <h2 className="text-2xl font-bold mb-6">Interview Pipeline</h2>
              <div className={`${cardBg} rounded-2xl shadow-sm border ${borderColor} overflow-hidden`}>
                <div className="overflow-x-auto">
                <table className="w-full text-left">
                  <thead>
                    <tr className={`${isDark ? 'bg-slate-950/50' : 'bg-slate-50'} text-[10px] font-bold ${textColorMuted} uppercase tracking-widest`}>
                      <th className="px-6 py-4">Candidate</th>
                      <th className="px-6 py-4">Status</th>
                      <th className="px-6 py-4">Schedule Interview</th>
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${borderColor}`}>
                    {evalResults?.candidates
                        .filter(c => shortlistedCandidateFilenames.includes(c.filename))
                        .map((candidate, idx) => (
                      <tr key={idx} className={`group hover:${isDark ? 'bg-slate-800/50' : 'bg-slate-50'} transition-colors`}>
                        <td className="px-6 py-5">
                            <p className={`text-sm font-bold ${text}`}>{candidate.name}</p>
                        </td>
                        <td className="px-6 py-5">
                            <span className="text-[10px] font-bold bg-indigo-100 text-indigo-700 px-2 py-1 rounded">Shortlisted</span>
                        </td>
                        <td className="px-6 py-5">
                            <input type="date" className={`${inputBg} border ${inputBorder} rounded-lg p-2 text-sm`}/>
                            <button className="ml-2 text-xs px-4 py-2 bg-indigo-600 text-white rounded-lg">Schedule</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
              </div>
          </Layout>
      )}


      {/* Candidate Details Modal */}
      <AnimatePresence>
        {selectedCandidate && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setSelectedCandidate(null)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ scale: 0.9, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.9, opacity: 0, y: 20 }}
              className="relative w-full max-w-2xl bg-white rounded-3xl shadow-2xl overflow-hidden overflow-y-auto max-h-[90vh]"
            >
              <div className="p-8 border-b border-slate-100 flex justify-between items-start">
                <div className="flex gap-4">
                  <div className="w-16 h-16 rounded-2xl bg-indigo-600 flex items-center justify-center text-white text-2xl font-bold">
                    {selectedCandidate.name.charAt(0)}
                  </div>
                  <div>
                    <h2 className="text-2xl font-bold text-slate-900">{selectedCandidate.name}</h2>
                    <p className="text-slate-500 font-medium flex items-center gap-1.5 text-sm mt-1">
                      <FileText className="w-4 h-4" /> {selectedCandidate.filename}
                    </p>
                    {selectedCandidate.email && (
                      <p className="text-slate-500 font-medium flex items-center gap-1.5 text-sm mt-1">
                        📧 {selectedCandidate.email}
                      </p>
                    )}
                    {selectedCandidate.phone && (
                      <p className="text-slate-500 font-medium flex items-center gap-1.5 text-sm mt-1">
                        📞 {selectedCandidate.phone}
                      </p>
                    )}
                  </div>
                </div>
                <button 
                  onClick={() => setSelectedCandidate(null)}
                  className="p-2 hover:bg-slate-100 rounded-full transition-colors"
                >
                  <X className="w-6 h-6 text-slate-400" />
                </button>
              </div>

              <div className="p-8 space-y-8">
                <div className="grid grid-cols-3 gap-4">
                  <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1">Match Score</p>
                    <p className="text-2xl font-black text-indigo-600">{selectedCandidate.score}%</p>
                  </div>
                  <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1">Verdict</p>
                    <p className="text-sm font-bold text-slate-700">{selectedCandidate.verdict}</p>
                  </div>
                  <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1">Status</p>
                    <span className="inline-flex items-center gap-1.5 text-[10px] font-bold text-emerald-600 uppercase">
                      <div className="w-2 h-2 bg-emerald-500 rounded-full" /> Shortlisted
                    </span>
                  </div>
                </div>

                <div>
                  <h4 className="text-xs font-black text-slate-900 uppercase tracking-widest mb-4">AI Recruiter Justification</h4>
                  <div className="bg-indigo-50/50 p-6 rounded-2xl border border-indigo-100 italic">
                    <p className="text-slate-700 text-sm leading-relaxed whitespace-pre-wrap leading-7 font-medium">
                      "{selectedCandidate.justification}"
                    </p>
                  </div>
                </div>

                <div>
                   <h4 className="text-xs font-black text-slate-900 uppercase tracking-widest mb-4">Education</h4>
                   <ul className="text-sm text-slate-700 list-disc pl-4 space-y-1">
                     {selectedCandidate.education.map((e: string, i: number) => <li key={i}>{e}</li>)}
                   </ul>
                </div>

                <div>
                   <h4 className="text-xs font-black text-slate-900 uppercase tracking-widest mb-4">Work Experience</h4>
                    <ul className="text-sm text-slate-700 list-disc pl-4 space-y-1">
                     {selectedCandidate.workExperience.map((e: string, i: number) => <li key={i}>{e}</li>)}
                   </ul>
                </div>

                {selectedCandidate.portfolioLink && (
                  <div>
                    <h4 className="text-xs font-black text-slate-900 uppercase tracking-widest mb-2">Portfolio</h4>
                    <a href={selectedCandidate.portfolioLink} target="_blank" rel="noopener noreferrer" className="text-sm text-blue-600 hover:underline">{selectedCandidate.portfolioLink}</a>
                  </div>
                )}
                
                {resumeUrl && (
                  <div>
                    <h4 className="text-xs font-black text-slate-900 uppercase tracking-widest mb-2">Resume Preview</h4>
                    <PDFPreview url={resumeUrl} />
                  </div>
                )}
                
                <div>
                    <h4 className="text-xs font-black text-slate-900 uppercase tracking-widest mb-2">Years of Experience</h4>
                    <p className="text-sm text-slate-700">{selectedCandidate.yearsOfExperience} years</p>
                </div>

                <div>
                  <h4 className="text-xs font-black text-slate-900 uppercase tracking-widest mb-4">Analyzed Skills & Competencies</h4>
                  <div className="flex flex-wrap gap-2">
                    {selectedCandidate.skills.map((skill, i) => (
                      <span key={i} className="px-4 py-2 bg-white border border-slate-200 rounded-xl text-xs font-bold text-slate-600 flex items-center gap-2">
                        <div className="w-1.5 h-1.5 bg-indigo-400 rounded-full" /> {skill}
                      </span>
                    ))}
                  </div>
                </div>
              </div>

              <div className="p-8 bg-slate-50 border-t border-slate-100 flex flex-col gap-4">
                <input 
                  type="text" 
                  placeholder="Enter Google Spreadsheet ID" 
                  value={spreadsheetIdInput} 
                  onChange={e => setSpreadsheetIdInput(e.target.value)}
                  className="w-full px-4 py-3 rounded-xl border border-slate-200 text-sm"
                />
                <button onClick={() => handleProceedToInterview(selectedCandidate)} className="flex-1 py-4 bg-indigo-600 text-white rounded-2xl font-bold text-sm shadow-xl shadow-indigo-100 hover:bg-indigo-700 transition-all flex items-center justify-center gap-2">
                   Proceed to Interview <ChevronRight className="w-4 h-4" />
                </button>
                <button className="flex-1 py-4 bg-white border border-slate-200 text-slate-700 rounded-2xl font-bold text-sm hover:bg-slate-100 transition-all">
                  Reject Candidate
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

    </div>
  );
}
