import React, { useEffect, useMemo, useRef, useState } from "react";
import { Plus, Trash2, Printer, Mail, MessageCircle, Upload, Copy, Settings, FileText, FileDown, Save } from "lucide-react";
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';

// --- UI COMPONENTS ---
const Button = ({ children, variant = "default", size = "default", className, ...props }) => {
  const variants = {
    default: "bg-slate-900 text-slate-50 hover:bg-slate-900/90",
    outline: "border border-slate-200 bg-white hover:bg-slate-100 hover:text-slate-900",
    ghost: "hover:bg-slate-100 hover:text-slate-900",
  };
  const sizes = {
    default: "h-10 px-4 py-2",
    sm: "h-9 rounded-md px-3",
    lg: "h-11 rounded-md px-8",
  };
  return (
    <button className={`inline-flex items-center justify-center rounded-md text-sm font-medium transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-slate-400 disabled:pointer-events-none disabled:opacity-50 ${variants[variant]} ${sizes[size]} ${className || ''}`} {...props}>
      {children}
    </button>
  );
};

const Input = ({ className, ...props }) => (
  <input className={`flex h-10 w-full rounded-md border border-slate-200 bg-white px-3 py-2 text-sm ring-offset-white file:border-0 file:bg-transparent file:text-sm file:font-medium placeholder:text-slate-500 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-slate-400 disabled:cursor-not-allowed disabled:opacity-50 ${className || ''}`} {...props} />
);

const Textarea = ({ className, ...props }) => (
  <textarea className={`flex min-h-[80px] w-full rounded-md border border-slate-200 bg-white px-3 py-2 text-sm ring-offset-white placeholder:text-slate-500 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-slate-400 disabled:cursor-not-allowed disabled:opacity-50 ${className || ''}`} {...props} />
);

const Label = ({ className, ...props }) => (
  <label className={`text-sm font-medium leading-none peer-disabled:cursor-not-allowed peer-disabled:opacity-70 ${className || ''}`} {...props} />
);

const Card = ({ className, ...props }) => (
  <div className={`rounded-xl border border-slate-200 bg-white text-slate-950 shadow-sm ${className || ''}`} {...props} />
);
const CardHeader = ({ className, ...props }) => <div className={`flex flex-col space-y-1.5 p-6 ${className || ''}`} {...props} />;
const CardTitle = ({ className, ...props }) => <h3 className={`text-2xl font-semibold leading-none tracking-tight ${className || ''}`} {...props} />;
const CardContent = ({ className, ...props }) => <div className={`p-6 pt-0 ${className || ''}`} {...props} />;

const Checkbox = ({ checked, onCheckedChange, className }) => (
  <input type="checkbox" checked={checked} onChange={(e) => onCheckedChange(e.target.checked)} className={`h-4 w-4 rounded border-slate-300 text-slate-900 focus:ring-slate-400 ${className || ''}`} />
);

const Switch = ({ checked, onCheckedChange, className }) => (
  <button type="button" role="switch" aria-checked={checked} onClick={() => onCheckedChange(!checked)} className={`peer inline-flex h-[24px] w-[44px] shrink-0 cursor-pointer items-center rounded-full border-2 border-transparent transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-slate-400 focus-visible:ring-offset-2 focus-visible:ring-offset-white disabled:cursor-not-allowed disabled:opacity-50 ${checked ? 'bg-slate-900' : 'bg-slate-200'} ${className || ''}`}>
    <span className={`pointer-events-none block h-5 w-5 rounded-full bg-white shadow-lg ring-0 transition-transform ${checked ? 'translate-x-5' : 'translate-x-0'}`} />
  </button>
);

const Badge = ({ variant = "default", className, ...props }) => {
  const variants = {
    default: "border-transparent bg-slate-900 text-slate-50 hover:bg-slate-900/80",
    secondary: "border-transparent bg-slate-100 text-slate-900 hover:bg-slate-100/80",
  };
  return <div className={`inline-flex items-center rounded-full border px-2.5 py-0.5 text-xs font-semibold transition-colors focus:outline-none focus:ring-2 focus:ring-slate-400 focus:ring-offset-2 ${variants[variant]} ${className || ''}`} {...props} />;
};

// Tabs Implementation using Context
const TabsContext = React.createContext({});
const Tabs = ({ value, onValueChange, children, className }) => (
  <TabsContext.Provider value={{ value, onValueChange }}>
    <div className={className || ''}>{children}</div>
  </TabsContext.Provider>
);
const TabsList = ({ children, className }) => (
  <div className={`inline-flex h-10 items-center justify-center rounded-md bg-slate-100 p-1 text-slate-500 ${className || ''}`}>{children}</div>
);
const TabsTrigger = ({ value, children, className }) => {
  const { value: selectedValue, onValueChange } = React.useContext(TabsContext);
  const isSelected = selectedValue === value;
  return (
    <button type="button" onClick={() => onValueChange(value)} className={`inline-flex items-center justify-center whitespace-nowrap rounded-sm px-3 py-1.5 text-sm font-medium ring-offset-white transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-slate-400 focus-visible:ring-offset-2 disabled:pointer-events-none disabled:opacity-50 ${isSelected ? 'bg-white text-slate-950 shadow-sm' : ''} ${className || ''}`}>
      {children}
    </button>
  );
};
const TabsContent = ({ value, children, className, ...props }) => {
  const { value: selectedValue } = React.useContext(TabsContext);
  if (selectedValue !== value) return null;
  return <div className={`mt-2 ring-offset-white focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-slate-400 focus-visible:ring-offset-2 ${className || ''}`} {...props}>{children}</div>;
};

// --- APP LOGIC & STATE ---
const STORAGE_KEYS = {
  settings: "gmds_settings_v1",
  sheets: "gmds_sheets_v1",
  squad: "gmds_squad_v1",
};

const NOTICE_TEXT = "After match please complete the Fulltime Scores via Fulltime or Match day App. The league reserves the right to request match cards for audit purposes at any time. Do not report half-time scores. Match officials are paid prior to the fixture commencing.";

const defaultTheme = {
  header: "#0f172a",
  background: "#f8fafc",
  button: "#dc2626",
};

const createId = () => {
  if (typeof crypto !== "undefined" && crypto.randomUUID) return crypto.randomUUID();
  return `id-${Date.now()}-${Math.random().toString(36).slice(2, 11)}`;
};

const blankPlayer = (index = 0) => ({
  id: createId(),
  shirtNumber: index + 1 <= 11 ? String(index + 1) : "",
  playerName: "",
  idChecked: false,
  goals: "",
  yellow: false,
  red: false,
});

const buildPlayers = (count = 11) => Array.from({ length: count }, (_, i) => blankPlayer(i));

const defaultSettings = {
  clubName: "",
  defaultManagerName: "",
  defaultLogo: "",
  theme: defaultTheme,
};

const defaultSquad = buildPlayers(11).map(p => ({ id: p.id, shirtNumber: p.shirtNumber, playerName: "" }));


const createNewSheet = (settings) => ({
  id: createId(),
  status: "draft",
  createdAt: new Date().toISOString(),
  updatedAt: new Date().toISOString(),
  league: "WJFL",
  date: new Date().toISOString().slice(0, 10),
  divisionOrCup: "",
  groundName: "",
  refereeName: "",
  leagueAssignedRef: false,
  refereeSignature: "",
  refereeSignedDate: "",
  refereeFeePaid: "",
  refereeReport: "",
  logo: settings?.defaultLogo || "",
  teamName: settings?.clubName || "",
  managerName: settings?.defaultManagerName || "",
  oppositionTeam: "",
  oppositionManager: "",
  oppositionIdChecksCompleted: false,
  oppositionIdChecksNotes: "",
  oppositionHaveMatchSheet: false,
  oppositionHaveMatchSheetNotes: "",
  scoreFor: "",
  scoreAgainst: "",
  penaltiesFor: "",
  penaltiesAgainst: "",
  players: buildPlayers(11),
  managerSignature: "",
  managerSignedDate: "",
  oppositionManagerSignature: "",
  oppositionManagerSignedDate: "",
});

function loadJson(key, fallback) {
  try {
    if (typeof window === "undefined") return fallback;
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

function saveJson(key, value) {
  try {
    if (typeof window === "undefined") return;
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Ignore local storage write errors
  }
}

function normalizeSheet(sheet, settings = defaultSettings) {
  const base = createNewSheet(settings);
  const migratedPlayers = Array.isArray(sheet?.players)
    ? sheet.players
    : Array.isArray(sheet?.homePlayers)
      ? sheet.homePlayers
      : buildPlayers(11);
  return {
    ...base,
    ...sheet,
    teamName: sheet?.teamName ?? sheet?.homeTeam ?? base.teamName,
    managerName: sheet?.managerName ?? sheet?.homeManager ?? base.managerName,
    oppositionTeam: sheet?.oppositionTeam ?? sheet?.awayTeam ?? base.oppositionTeam,
    oppositionManager: sheet?.oppositionManager ?? sheet?.awayManager ?? base.oppositionManager,
    scoreFor: sheet?.scoreFor ?? sheet?.scoreHome ?? base.scoreFor,
    scoreAgainst: sheet?.scoreAgainst ?? sheet?.scoreAway ?? base.scoreAgainst,
    penaltiesFor: sheet?.penaltiesFor ?? sheet?.penaltiesHome ?? base.penaltiesFor,
    penaltiesAgainst: sheet?.penaltiesAgainst ?? sheet?.penaltiesAway ?? base.penaltiesAgainst,
    managerSignature: sheet?.managerSignature ?? sheet?.homeManagerSignature ?? base.managerSignature,
    managerSignedDate: sheet?.managerSignedDate ?? sheet?.homeManagerSignedDate ?? base.managerSignedDate,
    oppositionManagerSignature: sheet?.oppositionManagerSignature ?? sheet?.awayManagerSignature ?? base.oppositionManagerSignature,
    oppositionManagerSignedDate: sheet?.oppositionManagerSignedDate ?? sheet?.awayManagerSignedDate ?? base.oppositionManagerSignedDate,
    players: migratedPlayers.map((player, index) => ({
      ...blankPlayer(index),
      ...player,
      id: player?.id || createId(),
    })),
  };
}

function formatDateTime(iso) {
  if (!iso) return "";
  return new Date(iso).toLocaleString();
}

function SignaturePad({ value, onChange, label }) {
  const canvasRef = useRef(null);
  const wrapperRef = useRef(null);
  const drawing = useRef(false);
  useEffect(() => {
    const canvas = canvasRef.current;
    const wrapper = wrapperRef.current;
    if (!canvas || !wrapper) return;
    const resize = () => {
      const ratio = window.devicePixelRatio || 1;
      const width = wrapper.clientWidth;
      const height = 160;
      const ctx = canvas.getContext("2d");
      const existing = canvas.toDataURL();
      
      canvas.width = width * ratio;
      canvas.height = height * ratio;
      canvas.style.width = `${width}px`;
      canvas.style.height = `${height}px`;
      
      ctx.scale(ratio, ratio);
      ctx.lineWidth = 2;
      ctx.lineCap = "round";
      ctx.strokeStyle = "#111827";
      
      if (value || existing !== "data:,") {
        const img = new Image();
        img.onload = () => ctx.drawImage(img, 0, 0, width, height);
        img.src = value || existing;
      }
    };
    resize();
    window.addEventListener("resize", resize);
    return () => window.removeEventListener("resize", resize);
  }, [value]);
  const getPos = (e) => {
    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    const point = e.touches?.[0] || e;
    return { x: point.clientX - rect.left, y: point.clientY - rect.top };
  };
  const start = (e) => {
    e.preventDefault();
    drawing.current = true;
    const ctx = canvasRef.current.getContext("2d");
    const { x, y } = getPos(e);
    ctx.beginPath();
    ctx.moveTo(x, y);
  };
  const move = (e) => {
    if (!drawing.current) return;
    e.preventDefault();
    const ctx = canvasRef.current.getContext("2d");
    const { x, y } = getPos(e);
    ctx.lineTo(x, y);
    ctx.stroke();
  };
  const end = () => {
    if (!drawing.current) return;
    drawing.current = false;
    onChange(canvasRef.current.toDataURL("image/png"));
  };
  const clear = () => {
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    onChange("");
  };
  return (
    <div className="space-y-2" ref={wrapperRef}>
      <div className="flex items-center justify-between gap-3">
        <Label>{label}</Label>
        <Button type="button" variant="outline" size="sm" onClick={clear}>Clear signature</Button>
      </div>
      <canvas
        ref={canvasRef}
        className="w-full rounded-xl border bg-slate-50 touch-none"
        onMouseDown={start}
        onMouseMove={move}
        onMouseUp={end}
        onMouseLeave={end}
        onTouchStart={start}
        onTouchMove={move}
        onTouchEnd={end}
      />
    </div>
  );
}

function PlayerTable({ title, players, onChange, theme, squad }) {
  const updatePlayer = (idx, key, value) => {
    const next = [...players];
    next[idx] = { ...next[idx], [key]: value };
    onChange(next);
  };
  const addPlayer = () => onChange([...players, blankPlayer(players.length)]);
  const removePlayer = (idx) => {
    if (players.length <= 3) return;
    onChange(players.filter((_, i) => i !== idx));
  };

  const allChecked = players.every((p) => !!p.idChecked);
  const toggleAll = (checked) => {
    onChange(players.map((p) => ({ ...p, idChecked: checked })));
  };

  const importSquad = () => {
    if (!squad || !squad.length) return;
    const imported = squad
      .filter(sp => sp.playerName.trim() !== "")
      .map((sp, idx) => ({
        ...blankPlayer(idx),
        shirtNumber: sp.shirtNumber,
        playerName: sp.playerName
      }));
    if (imported.length > 0) {
      onChange(imported);
    }
  };

  return (
    <Card className="rounded-2xl shadow-sm">
      <CardHeader>
        <CardTitle className="flex items-center justify-between">
          <div className="flex items-center gap-4">
            <span>{title}</span>
            {squad && squad.some(p => p.playerName.trim() !== "") && (
              <Button type="button" size="sm" variant="outline" onClick={importSquad}>
                Import Squad
              </Button>
            )}
          </div>
          <Button type="button" size="sm" onClick={addPlayer} style={{ backgroundColor: theme.button }} className="text-white">
            <Plus className="mr-2 h-4 w-4" />Add player
          </Button>
        </CardTitle>
      </CardHeader>
      <CardContent className="overflow-x-auto">
        <table className="w-full min-w-[900px] border-collapse text-sm">
          <thead>
            <tr className="bg-slate-100 text-left">
              <th className="border p-2">Shirt Number</th>
              <th className="border p-2">Player Name</th>
              <th className="border p-2">
                <div className="flex items-center gap-2">
                  <Checkbox checked={allChecked} onCheckedChange={toggleAll} />
                  <span>ID Checked by Opposition</span>
                </div>
              </th>
              <th className="border p-2">Goals</th>
              <th className="border p-2">Yellow</th>
              <th className="border p-2">Red</th>
              <th className="border p-2">Action</th>
            </tr>
          </thead>
          <tbody>
            {players.map((player, idx) => (
              <tr key={player.id}>
                <td className="border p-2"><Input value={player.shirtNumber} onChange={(e) => updatePlayer(idx, "shirtNumber", e.target.value)} /></td>
                <td className="border p-2"><Input value={player.playerName} onChange={(e) => updatePlayer(idx, "playerName", e.target.value)} /></td>
                <td className="border p-2">
                  <div className="flex items-center gap-2">
                    <Switch checked={!!player.idChecked} onCheckedChange={(v) => updatePlayer(idx, "idChecked", v)} />
                    <span>{player.idChecked ? "Yes" : "No"}</span>
                  </div>
                </td>
                <td className="border p-2"><Input value={player.goals} onChange={(e) => updatePlayer(idx, "goals", e.target.value)} /></td>
                <td className="border p-2 text-center"><Checkbox checked={!!player.yellow} onCheckedChange={(v) => updatePlayer(idx, "yellow", !!v)} /></td>
                <td className="border p-2 text-center"><Checkbox checked={!!player.red} onCheckedChange={(v) => updatePlayer(idx, "red", !!v)} /></td>
                <td className="border p-2">
                  <Button type="button" variant="outline" size="sm" onClick={() => removePlayer(idx)}>
                    <Trash2 className="h-4 w-4" />
                  </Button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </CardContent>
    </Card>
  );
}

export default function App() {
  const [settings, setSettings] = useState(defaultSettings);
  const [sheets, setSheets] = useState([]);
  const [squad, setSquad] = useState([]);
  const [activeId, setActiveId] = useState("");
  const [tab, setTab] = useState("editor");
  const [deferredPrompt, setDeferredPrompt] = useState(null);

  useEffect(() => {
    const handler = (e) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };
    window.addEventListener("beforeinstallprompt", handler);
    return () => window.removeEventListener("beforeinstallprompt", handler);
  }, []);

  const installApp = async () => {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    await deferredPrompt.userChoice;
    setDeferredPrompt(null);
  };

  useEffect(() => {
    const savedSettings = loadJson(STORAGE_KEYS.settings, defaultSettings);
    const mergedSettings = { ...defaultSettings, ...savedSettings, theme: { ...defaultTheme, ...(savedSettings.theme || {}) } };
    const savedSheets = loadJson(STORAGE_KEYS.sheets, []);
    const savedSquad = loadJson(STORAGE_KEYS.squad, defaultSquad);
    
    const normalizedSheets = Array.isArray(savedSheets)
      ? savedSheets.map((sheet) => normalizeSheet(sheet, mergedSettings))
      : [];
    
    setSettings(mergedSettings);
    setSquad(savedSquad);

    if (normalizedSheets.length) {
      setSheets(normalizedSheets);
      setActiveId(normalizedSheets[0].id);
    } else {
      const first = createNewSheet(mergedSettings);
      setSheets([first]);
      setActiveId(first.id);
    }
  }, []);

  useEffect(() => {
    saveJson(STORAGE_KEYS.settings, settings);
  }, [settings]);

  useEffect(() => {
    if (sheets.length) saveJson(STORAGE_KEYS.sheets, sheets);
  }, [sheets]);

  useEffect(() => {
    saveJson(STORAGE_KEYS.squad, squad);
  }, [squad]);


  const activeSheet = useMemo(() => {
    if (!sheets.length) return null;
    return sheets.find((s) => s.id === activeId) || sheets[0] || null;
  }, [sheets, activeId]);

  const updateSheet = (patch) => {
    const targetId = activeId || activeSheet?.id;
    if (!targetId) return;
    setSheets((prev) => prev.map((sheet) => sheet.id === targetId ? { ...sheet, ...patch, updatedAt: new Date().toISOString() } : sheet));
  };

  const validateForCompletion = (sheet) => {
    const missing = [];
    if (!sheet.refereeSignature) missing.push("Referee signature");
    if (!sheet.managerSignature) missing.push("Team manager signature");
    if (!sheet.oppositionManagerSignature) missing.push("Opposition manager signature");
    return missing;
  };

  const newSheet = () => {
    const sheet = createNewSheet(settings);
    setSheets((prev) => [sheet, ...prev]);
    setActiveId(sheet.id);
    setTab("editor");
  };

  const duplicateSheet = () => {
    if (!activeSheet) return;
    const copy = {
      ...activeSheet,
      id: createId(),
      status: "draft",
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    setSheets((prev) => [copy, ...prev]);
    setActiveId(copy.id);
    setTab("editor");
  };

  const deleteSheet = (id) => {
    const next = sheets.filter((s) => s.id !== id);
    if (!next.length) {
      const fresh = createNewSheet(settings);
      setSheets([fresh]);
      setActiveId(fresh.id);
      setTab("editor");
      return;
    }
    setSheets(next);
    if (id === activeId) {
      setActiveId(next[0].id);
      setTab("editor");
    }
  };

  const saveDraft = () => updateSheet({ status: "draft" });
  const markComplete = () => {
    const missing = validateForCompletion(activeSheet);
    if (missing.length) {
      alert(`This sheet cannot be completed yet. Missing: ${missing.join(", ")}`);
      return;    }
    updateSheet({ status: "complete" });
    alert("Sheet marked as complete.");
  };

  const exportAllXlsx = () => {
    const data = [];
    sheets.forEach(sheet => {
      // Add sheet header info
      data.push({
        'Match ID': sheet.id,
        'Team': sheet.teamName,
        'Opposition': sheet.oppositionTeam,
        'League': sheet.league,
        'Date': sheet.date,
        'Score': `${sheet.scoreFor} - ${sheet.scoreAgainst}`,
        'Type': 'MATCH_HEADER'
      });
      
      // Add player details
      sheet.players.forEach(p => {
        data.push({
          'Match ID': '',
          'Team': sheet.teamName,
          'Opposition': sheet.oppositionTeam,
          'League': sheet.league,
          'Date': sheet.date,
          'Score': '',
          'Type': 'PLAYER',
          'Shirt #': p.shirtNumber,
          'Player Name': p.playerName,
          'ID Checked': p.idChecked ? 'Yes' : 'No',
          'Goals': p.goals,
          'Yellow': p.yellow ? '1' : '0',
          'Red': p.red ? '1' : '0'
        });
      });
      
      // Blank row between sheets
      data.push({});
    });
    
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Match Sheets");
    XLSX.writeFile(wb, "match-day-sheets.xlsx");
  };

  const exportAllHtml = () => {
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>All Match Sheets</title>
        <style>
          body { font-family: sans-serif; padding: 20px; background: #f8fafc; }
          .sheet { background: white; border: 1px solid #e2e8f0; border-radius: 12px; padding: 20px; margin-bottom: 40px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); page-break-after: always; }
          h1 { color: #0f172a; border-bottom: 2px solid #0f172a; padding-bottom: 10px; }
          .meta { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 20px; }
          table { width: 100%; border-collapse: collapse; margin-top: 10px; }
          th, td { border: 1px solid #e2e8f0; padding: 8px; text-align: left; }
          th { background: #f1f5f9; }
          .notice { color: #dc2626; font-weight: bold; margin-top: 20px; font-size: 0.9em; }
        </style>
      </head>
      <body>
        ${sheets.map(sheet => `
          <div class="sheet">
            <h1>Match Day Sheet: ${sheet.teamName} vs ${sheet.oppositionTeam}</h1>
            <div class="meta">
              <div><strong>League:</strong> ${sheet.league}</div>
              <div><strong>Date:</strong> ${sheet.date}</div>
              <div><strong>Division/Cup:</strong> ${sheet.divisionOrCup}</div>
              <div><strong>Ground:</strong> ${sheet.groundName}</div>
              <div><strong>Referee:</strong> ${sheet.refereeName}</div>
              <div><strong>Score:</strong> ${sheet.scoreFor} - ${sheet.scoreAgainst}</div>
            </div>
            <h2>Lineup</h2>
            <table>
              <thead>
                <tr>
                  <th>#</th>
                  <th>Player Name</th>
                  <th>ID Checked</th>
                  <th>Goals</th>
                  <th>YEL</th>
                  <th>RED</th>
                </tr>
              </thead>
              <tbody>
                ${sheet.players.map(p => `
                  <tr>
                    <td>${p.shirtNumber}</td>
                    <td>${p.playerName}</td>
                    <td>${p.idChecked ? 'Yes' : 'No'}</td>
                    <td>${p.goals}</td>
                    <td>${p.yellow ? 'Y' : ''}</td>
                    <td>${p.red ? 'R' : ''}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
            <div class="notice">${NOTICE_TEXT}</div>
          </div>
        `).join('')}
      </body>
      </html>
    `;
    const blob = new Blob([htmlContent], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `all-match-sheets.html`;
    a.click();
    URL.revokeObjectURL(url);
  };


  const printSheet = () => window.print();
  
  const emailSheet = (sheet) => {
    // 1. Open Email App
    const subject = encodeURIComponent(`Match Day Sheet - ${sheet.teamName} vs ${sheet.oppositionTeam}`);
    const body = encodeURIComponent(`Match Day Sheet\n\nLeague: ${sheet.league}\nDate: ${sheet.date}\nTeam: ${sheet.teamName}\nOpposition: ${sheet.oppositionTeam}\nScore: ${sheet.scoreFor} - ${sheet.scoreAgainst}\n\nPlease attach the PDF match sheet (you can generate this by clicking 'Print / PDF' on the website and choosing 'Save as PDF').`);
    
    window.location.href = `mailto:?subject=${subject}&body=${body}`;
    
    // 2. Alert the user with instructions
    alert("I've opened your email app with the match details filled in. To include the full sheet, click the 'Print / PDF' button on the app, choose 'Save as PDF', and then attach that file to your email!");
  };

  const whatsappSheet = (sheet) => {
    const text = encodeURIComponent(`Match Day Sheet\n${sheet.teamName} vs ${sheet.oppositionTeam}\n${sheet.date}\nScore: ${sheet.scoreFor} - ${sheet.scoreAgainst}`);
    window.open(`https://wa.me/?text=${text}`, "_blank");
  };

  const onLogoUpload = (file, target = "default") => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || "");
      if (target === "default") {
        setSettings((prev) => ({ ...prev, defaultLogo: result }));
      } else {
        updateSheet({ logo: result });
      }
    };
    reader.readAsDataURL(file);
  };

  const exportSquadJson = () => {
    const blob = new Blob([JSON.stringify(squad, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `my-squad-backup.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const importSquadJson = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const imported = JSON.parse(e.target.result);
        if (Array.isArray(imported)) {
          setSquad(imported);
          alert("Squad imported successfully!");
        }
      } catch (err) {
        alert("Error importing squad. Please make sure it's a valid backup file.");
      }
    };
    reader.readAsText(file);
  };

  if (!activeSheet) return <div className="p-4">Loading sheet...</div>;

  const themeVars = {
    backgroundColor: settings.theme.background,
    minHeight: "100vh",
  };

  return (
    <div style={themeVars} className="p-4 md:p-6 print:p-0">
      <div className="mx-auto max-w-7xl space-y-6">
        <div className="flex flex-col gap-4 rounded-3xl p-5 text-white shadow-lg print:rounded-none" style={{ backgroundColor: settings.theme.header }}>
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div className="flex items-center gap-4">
              {settings.defaultLogo ? <img src={settings.defaultLogo} alt="Club logo" className="h-16 w-16 rounded-2xl border bg-white object-contain p-2" /> : null}
              <div>
                <h1 className="text-2xl font-bold">Grassroots Football Match Day Sheets</h1>
                <p className="text-sm text-slate-200">Offline-friendly local app for multiple clubs, teams and coaches</p>
              </div>
            </div>
            <div className="flex flex-wrap gap-2 print:hidden">
              {deferredPrompt && (
                <Button onClick={installApp} className="bg-green-600 text-white border-green-700 hover:bg-green-700 animate-pulse">
                  <Upload className="mr-2 h-4 w-4" /> Install App
                </Button>
              )}
              <Button onClick={newSheet} style={{ backgroundColor: settings.theme.button }} className="text-white"><Plus className="mr-2 h-4 w-4" />New sheet</Button>
              <Button variant="outline" className="border-white bg-white text-slate-900 hover:bg-slate-100 hover:text-slate-900" onClick={duplicateSheet}><Copy className="mr-2 h-4 w-4" />Duplicate</Button>
              <Button variant="outline" className="border-white bg-white text-slate-900 hover:bg-slate-100 hover:text-slate-900" onClick={saveDraft}><Save className="mr-2 h-4 w-4" />Save draft</Button>
              <Button variant="outline" className="border-white bg-white text-slate-900 hover:bg-slate-100 hover:text-slate-900" onClick={markComplete}>Complete</Button>
            </div>
          </div>
        </div>
        <Tabs value={tab} onValueChange={setTab} className="space-y-6">
          <TabsList className="grid w-full grid-cols-3 print:hidden">
            <TabsTrigger value="sheets">Match sheets</TabsTrigger>
            <TabsTrigger value="editor">Active sheet</TabsTrigger>
            <TabsTrigger value="settings">Settings</TabsTrigger>
          </TabsList>
          <TabsContent value="sheets" className="space-y-4 print:hidden">
            <Card className="rounded-2xl shadow-sm">
              <CardHeader className="flex flex-row items-center justify-between">
                <CardTitle>Saved match sheets</CardTitle>
                <div className="flex gap-2">
                  <Button variant="outline" onClick={exportAllXlsx}><FileDown className="mr-2 h-4 w-4" />Export all (Excel)</Button>
                  <Button variant="outline" onClick={exportAllHtml}><FileDown className="mr-2 h-4 w-4" />Export all (HTML)</Button>
                </div>
              </CardHeader>
              <CardContent className="space-y-3">
                {sheets.map((sheet) => (
                  <div key={sheet.id} className={`flex flex-col gap-3 rounded-2xl border p-4 md:flex-row md:items-center md:justify-between ${sheet.id === activeId ? "border-slate-900 bg-slate-50/50" : "border-slate-200"}`}>
                    <div>
                      <div className="flex flex-wrap items-center gap-2">
                        <button className="text-left text-lg font-semibold underline-offset-4 hover:underline" onClick={() => { setActiveId(sheet.id); setTab("editor"); }}>
                          {sheet.teamName || "Team"} vs {sheet.oppositionTeam || "Opposition"}
                        </button>
                        <Badge variant={sheet.status === "complete" ? "default" : "secondary"}>{sheet.status}</Badge>
                      </div>
                      <p className="text-sm text-slate-600">{sheet.league} • {sheet.date || "No date"} • Updated {formatDateTime(sheet.updatedAt)}</p>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      <Button variant="outline" size="sm" onClick={() => { setActiveId(sheet.id); setTab("editor"); }}><FileText className="mr-2 h-4 w-4" />Open</Button>
                      <Button variant="outline" size="sm" onClick={() => emailSheet(sheet)}><Mail className="mr-2 h-4 w-4" />Email</Button>
                      <Button variant="outline" size="sm" onClick={() => whatsappSheet(sheet)}><MessageCircle className="mr-2 h-4 w-4" />WhatsApp</Button>
                      <Button variant="outline" size="sm" onClick={() => deleteSheet(sheet.id)}><Trash2 className="mr-2 h-4 w-4" />Delete</Button>
                    </div>
                  </div>
                ))}
              </CardContent>
            </Card>
          </TabsContent>
          <TabsContent value="editor" className="space-y-6" id="match-editor-content">
            <Card className="rounded-2xl shadow-sm">
              <CardHeader className="flex flex-row items-center justify-between">
                <CardTitle>Match details</CardTitle>
                <div className="flex flex-wrap gap-2 print:hidden">
                  <Button variant="outline" onClick={printSheet}><Printer className="mr-2 h-4 w-4" />Print / PDF</Button>
                  <Button variant="outline" onClick={() => emailSheet(activeSheet)}><Mail className="mr-2 h-4 w-4" />Email</Button>
                  <Button variant="outline" onClick={() => whatsappSheet(activeSheet)}><MessageCircle className="mr-2 h-4 w-4" />WhatsApp</Button>
                </div>
              </CardHeader>
              <CardContent className="grid gap-4 md:grid-cols-2">
                <div className="space-y-2">
                  <Label>League</Label>
                  <Input value={activeSheet.league} onChange={(e) => updateSheet({ league: e.target.value })} placeholder="Enter league name (e.g. WJFL)" />
                </div>
                <div className="space-y-2">
                  <Label>Date</Label>
                  <Input type="date" value={activeSheet.date} onChange={(e) => updateSheet({ date: e.target.value })} />
                </div>
                <div className="space-y-2">
                  <Label>Division or Cup selection</Label>
                  <Input value={activeSheet.divisionOrCup} onChange={(e) => updateSheet({ divisionOrCup: e.target.value })} placeholder="Enter division or cup" />
                </div>
                <div className="space-y-2">
                  <Label>Ground name</Label>
                  <Input value={activeSheet.groundName} onChange={(e) => updateSheet({ groundName: e.target.value })} />
                </div>
                <div className="space-y-2">
                  <Label>Referee name</Label>
                  <Input value={activeSheet.refereeName} onChange={(e) => updateSheet({ refereeName: e.target.value })} />
                </div>
                <div className="space-y-2">
                  <Label>League assigned ref</Label>
                  <div className="flex items-center gap-3 rounded-xl border p-3">
                    <Switch checked={!!activeSheet.leagueAssignedRef} onCheckedChange={(v) => updateSheet({ leagueAssignedRef: v })} />
                    <span>{activeSheet.leagueAssignedRef ? "Yes" : "No"}</span>
                  </div>
                </div>
                <div className="space-y-2">
                  <Label>Team name</Label>
                  <Input value={activeSheet.teamName} onChange={(e) => updateSheet({ teamName: e.target.value })} />
                </div>
                <div className="space-y-2">
                  <Label>Opposition team</Label>
                  <Input value={activeSheet.oppositionTeam} onChange={(e) => updateSheet({ oppositionTeam: e.target.value })} />
                </div>
                <div className="space-y-2">
                  <Label>Manager for team</Label>
                  <Input value={activeSheet.managerName} onChange={(e) => updateSheet({ managerName: e.target.value })} />
                </div>
                <div className="space-y-2">
                  <Label>Manager for opposition</Label>
                  <Input value={activeSheet.oppositionManager} onChange={(e) => updateSheet({ oppositionManager: e.target.value })} />
                </div>
                <div className="space-y-2">
                  <Label>Opposition ID checks completed</Label>
                  <div className="flex items-center gap-3 rounded-xl border p-3">
                    <Switch checked={!!activeSheet.oppositionIdChecksCompleted} onCheckedChange={(v) => updateSheet({ oppositionIdChecksCompleted: v })} />
                    <span>{activeSheet.oppositionIdChecksCompleted ? "Yes" : "No"}</span>
                  </div>
                </div>
                <div className="space-y-2">
                  <Label>Notes</Label>
                  <Textarea value={activeSheet.oppositionIdChecksNotes} onChange={(e) => updateSheet({ oppositionIdChecksNotes: e.target.value })} />
                </div>
                <div className="space-y-2">
                  <Label>Opposition have match sheet</Label>
                  <div className="flex items-center gap-3 rounded-xl border p-3">
                    <Switch checked={!!activeSheet.oppositionHaveMatchSheet} onCheckedChange={(v) => updateSheet({ oppositionHaveMatchSheet: v })} />
                    <span>{activeSheet.oppositionHaveMatchSheet ? "Yes" : "No"}</span>
                  </div>
                </div>
                <div className="space-y-2">
                  <Label>Notes</Label>
                  <Textarea value={activeSheet.oppositionHaveMatchSheetNotes} onChange={(e) => updateSheet({ oppositionHaveMatchSheetNotes: e.target.value })} />
                </div>
                <div className="space-y-2">
                  <Label>Score - Team</Label>
                  <Input value={activeSheet.scoreFor} onChange={(e) => updateSheet({ scoreFor: e.target.value })} />
                </div>
                <div className="space-y-2">
                  <Label>Score - Opposition</Label>
                  <Input value={activeSheet.scoreAgainst} onChange={(e) => updateSheet({ scoreAgainst: e.target.value })} />
                </div>
                <div className="space-y-2">
                  <Label>Penalties (Cup Game Only) - Team</Label>
                  <Input value={activeSheet.penaltiesFor} onChange={(e) => updateSheet({ penaltiesFor: e.target.value })} />
                </div>
                <div className="space-y-2">
                  <Label>Penalties (Cup Game Only) - Opposition</Label>
                  <Input value={activeSheet.penaltiesAgainst} onChange={(e) => updateSheet({ penaltiesAgainst: e.target.value })} />
                </div>
              </CardContent>
            </Card>
            <PlayerTable title="Team details" players={activeSheet.players} onChange={(players) => updateSheet({ players })} theme={settings.theme} squad={squad} />
            <Card className="rounded-2xl shadow-sm">
              <CardHeader>
                <CardTitle>Referee section</CardTitle>
              </CardHeader>
              <CardContent className="grid gap-4 md:grid-cols-2">
                <div className="space-y-2 md:col-span-2">
                  <Label>Referee report</Label>
                  <Textarea value={activeSheet.refereeReport} onChange={(e) => updateSheet({ refereeReport: e.target.value })} />
                </div>
                <div className="space-y-2">
                  <Label>Fee paid £</Label>
                  <Input value={activeSheet.refereeFeePaid} onChange={(e) => updateSheet({ refereeFeePaid: e.target.value })} />
                </div>
                <div className="space-y-2">
                  <Label>Referee signed date</Label>
                  <Input type="date" value={activeSheet.refereeSignedDate} onChange={(e) => updateSheet({ refereeSignedDate: e.target.value })} />
                </div>
                <div className="md:col-span-2">
                  <SignaturePad label="Referee signature" value={activeSheet.refereeSignature} onChange={(v) => updateSheet({ refereeSignature: v })} />
                </div>
              </CardContent>
            </Card>
            <Card className="rounded-2xl shadow-sm">
              <CardHeader>
                <CardTitle className="text-red-600">Important notice</CardTitle>
              </CardHeader>
              <CardContent>
                <p className="font-medium text-red-600">{NOTICE_TEXT}</p>
              </CardContent>
            </Card>
            <Card className="rounded-2xl shadow-sm">
              <CardHeader>
                <CardTitle>Manager signatures</CardTitle>
              </CardHeader>
              <CardContent className="grid gap-6 md:grid-cols-2">
                <div className="space-y-4">
                  <div className="space-y-2">
                    <SignaturePad label="Team manager signature" value={activeSheet.managerSignature} onChange={(v) => updateSheet({ managerSignature: v })} />
                  </div>
                  <div className="space-y-2">
                    <Label>Date</Label>
                    <Input type="date" value={activeSheet.managerSignedDate} onChange={(e) => updateSheet({ managerSignedDate: e.target.value })} />
                  </div>
                </div>
                <div className="space-y-4">
                  <div className="space-y-2">
                    <SignaturePad label="Opposition manager signature" value={activeSheet.oppositionManagerSignature} onChange={(v) => updateSheet({ oppositionManagerSignature: v })} />
                  </div>
                  <div className="space-y-2">
                    <Label>Date</Label>
                    <Input type="date" value={activeSheet.oppositionManagerSignedDate} onChange={(e) => updateSheet({ oppositionManagerSignedDate: e.target.value })} />
                  </div>
                </div>
              </CardContent>
            </Card>
          </TabsContent>
          <TabsContent value="settings" className="space-y-6 print:hidden">
            <Card className="rounded-2xl shadow-sm">
              <CardHeader>
                <CardTitle className="flex items-center gap-2"><Settings className="h-5 w-5" />Default app settings</CardTitle>
              </CardHeader>
              <CardContent className="grid gap-4 md:grid-cols-2">
                <div className="space-y-2">
                  <Label>Default club name</Label>
                  <Input value={settings.clubName} onChange={(e) => setSettings((prev) => ({ ...prev, clubName: e.target.value }))} />
                </div>
                <div className="space-y-2">
                  <Label>Default manager name</Label>
                  <Input value={settings.defaultManagerName} onChange={(e) => setSettings((prev) => ({ ...prev, defaultManagerName: e.target.value }))} />
                </div>
                <div className="space-y-2 md:col-span-2">
                  <Label>Default club logo</Label>
                  <div className="flex flex-wrap items-center gap-3 rounded-xl border p-3">
                    <Input type="file" accept="image/*" className="max-w-sm" onChange={(e) => onLogoUpload(e.target.files?.[0], "default")} />
                    {settings.defaultLogo ? <img src={settings.defaultLogo} alt="Default logo" className="h-16 w-16 rounded-xl border bg-white object-contain p-1" /> : null}
                  </div>
                </div>
                <div className="space-y-2">
                  <Label>Header colour</Label>
                  <Input type="color" className="p-1 h-12" value={settings.theme.header} onChange={(e) => setSettings((prev) => ({ ...prev, theme: { ...prev.theme, header: e.target.value } }))} />
                </div>
                <div className="space-y-2">
                  <Label>Background colour</Label>
                  <Input type="color" className="p-1 h-12" value={settings.theme.background} onChange={(e) => setSettings((prev) => ({ ...prev, theme: { ...prev.theme, background: e.target.value } }))} />
                </div>
                <div className="space-y-2">
                  <Label>Button colour</Label>
                  <Input type="color" className="p-1 h-12" value={settings.theme.button} onChange={(e) => setSettings((prev) => ({ ...prev, theme: { ...prev.theme, button: e.target.value } }))} />
                </div>
              </CardContent>
            </Card>

            <Card className="rounded-2xl shadow-sm">
              <CardHeader>
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div>
                    <CardTitle className="flex items-center gap-2">
                      <Plus className="h-5 w-5" />
                      My Squad (Defaults)
                    </CardTitle>
                    <p className="text-sm text-slate-500">Enter your regular team players here. Use the "Import Squad" button on a match sheet to fill them in instantly.</p>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    <Button variant="outline" size="sm" onClick={exportSquadJson}>
                      <FileDown className="mr-2 h-4 w-4" /> Export Squad File
                    </Button>
                    <div className="relative">
                      <input
                        type="file"
                        accept=".json"
                        className="absolute inset-0 opacity-0 cursor-pointer"
                        onChange={(e) => importSquadJson(e.target.files?.[0])}
                      />
                      <Button variant="outline" size="sm">
                        <Upload className="mr-2 h-4 w-4" /> Import Squad File
                      </Button>
                    </div>
                  </div>
                </div>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-sm">
                    <thead>
                      <tr className="bg-slate-100 text-left">
                        <th className="border p-2 w-24">Shirt #</th>
                        <th className="border p-2">Player Name</th>
                        <th className="border p-2 w-20">Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {squad.map((player, idx) => (
                        <tr key={player.id}>
                          <td className="border p-2">
                            <Input
                              value={player.shirtNumber}
                              onChange={(e) => {
                                const next = [...squad];
                                next[idx].shirtNumber = e.target.value;
                                setSquad(next);
                              }}
                            />
                          </td>
                          <td className="border p-2">
                            <Input
                              value={player.playerName}
                              onChange={(e) => {
                                const next = [...squad];
                                next[idx].playerName = e.target.value;
                                setSquad(next);
                              }}
                            />
                          </td>
                          <td className="border p-2 text-center">
                            <Button
                              variant="outline"
                              size="sm"
                              onClick={() => setSquad(squad.filter((_, i) => i !== idx))}
                            >
                              <Trash2 className="h-4 w-4" />
                            </Button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <Button
                  onClick={() => setSquad([...squad, { id: createId(), shirtNumber: "", playerName: "" }])}
                  variant="outline"
                  className="w-full"
                >
                  <Plus className="mr-2 h-4 w-4" /> Add Player to Squad
                </Button>
              </CardContent>
            </Card>

            <Card className="rounded-2xl shadow-sm">
              <CardHeader>
                <CardTitle>How to install on mobile</CardTitle>
              </CardHeader>
              <CardContent className="space-y-4 text-sm text-slate-700">
                <div>
                  <h3 className="font-bold text-slate-900 mb-1">iPhone (Safari)</h3>
                  <p>1. Tap the <strong>Share</strong> button (the square with an arrow) at the bottom.</p>
                  <p>2. Scroll down and tap <strong>'Add to Home Screen'</strong>.</p>
                </div>
                <div>
                  <h3 className="font-bold text-slate-900 mb-1">Android (Chrome)</h3>
                  <p>1. Look for the green <strong>'Install App'</strong> button at the top of this page.</p>
                  <p>2. If it's not there, tap the <strong>three dots</strong> in the corner and select <strong>'Install app'</strong>.</p>
                </div>
              </CardContent>
            </Card>

            <Card className="rounded-2xl shadow-sm">
              <CardHeader>
                <CardTitle>How this version works</CardTitle>
              </CardHeader>
              <CardContent className="space-y-2 text-sm text-slate-700">
                <p>• Saves everything locally in the browser so it works without login.</p>
                <p>• Drafts can be saved before signatures are added.</p>
                <p>• A sheet can only be marked complete when referee, team manager and opposition manager signatures are present.</p>
                <p>• Print can be used to create a PDF on most devices and browsers.</p>
                <p>• Email and WhatsApp use the device's own apps or web handling.</p>
                <p>• Export all currently saves a JSON backup of all sheets for portability.</p>
              </CardContent>
            </Card>
          </TabsContent>
        </Tabs>
      </div>
    </div>
  );
}
