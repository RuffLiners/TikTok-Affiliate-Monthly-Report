import React, { useState, useRef, useMemo, useEffect } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";

// ─── TYPES ───────────────────────────────────────────────────────────────────

interface VideoRow {
  id: string;
  source: string;
  creator: string;
  videoId: string;
  videoLink: string;
  revenue: number;
  itemsSold: number;
  views: number;
  likes: number;
  comments: number;
  description: string;
  hashtags: string;
  product: string;
  datePosted: string;
  audioHook: string;
  textHook: string;
  sellingPoints: string;
  visualHook: string;
  videoLength: string;
  cta: string;
  keyIdea: string;
  transcript: string;
  rank: number;
}

interface CreatorSummary {
  creator: string;
  totalVideos: number;
  unitsSold: number;
  gmv: number;
  videosWithGmv: number;
  videosLastMonth: number;
  videosWithGmvLastMonth: number;
  top3: VideoRow[];
}

interface HookSummary {
  hookText: string;
  totalVideos: number;
  totalViews: number;
  totalUnitsSold: number;
  totalGmv: number;
  videosWithGmv: number;
  topVideos: VideoRow[];
}

interface SellingPointSummary {
  point: string;
  product: string;
  totalVideos: number;
  totalViews: number;
  totalUnitsSold: number;
  totalGmv: number;
  videosWithGmv: number;
  topVideos: VideoRow[];
}

interface Override {
  audioHook?: string;
  visualHook?: string;
  textHook?: string;
  videoLength?: string;
  cta?: string;
  sellingPoints?: string;
  keyIdea?: string;
}

interface EditDraft {
  audioHook: string;
  visualHook: string;
  textHook: string;
  videoLength: string;
  cta: string;
  sellingPoints: string;
  keyIdea: string;
}

interface VideoCardProps {
  r: VideoRow;
  showFilter: boolean;
  hiddenIds: Set<string>;
  editingId: string | null;
  adminMode: boolean;
  transcriptOpen: Set<string>;
  visualHookOptions: string[];
  toggleHide: (videoId: string) => void;
  cancelEdit: () => void;
  openEdit: (r: VideoRow) => void;
  saveEdit: (r: VideoRow, draft: EditDraft) => void;
  toggleTranscript: (id: string) => void;
  onAddVisualHookOption: (v: string) => void;
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────

const C   = (s: unknown): string => typeof s === "string" ? s.replace(/\*\*/g,"").replace(/\[object Object\]/g,"").trim() : "";
const pts = (s: string): string[] => C(s).split("|").map((p: string) => p.trim()).filter(Boolean);
const hks = (s: string): string[] => pts(s).slice(0,3);
const lbl = (h: string): string => { const i=h.indexOf(":"); return (i>0&&i<65) ? h.slice(i+1).replace(/^"+|"+$/g,"").trim() : h.replace(/^"+|"+$/g,"").trim(); };
const f$  = (v: number): string => "$"+Number(v||0).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
const fN  = (v: number): string => Number(v||0).toLocaleString("en-US");
const fK  = (v: number): string => { const n=Number(v||0); return n>=1e6?(n/1e6).toFixed(1)+"M":n>=1e3?(n/1e3).toFixed(1)+"K":String(Math.round(n)); };

// ─── CSV PARSER ───────────────────────────────────────────────────────────────

// parseCSV always stores raw CSV values — overrides are applied separately so
// tiktok_reports stays clean and tiktok_overrides is the sole source for edits.
const parseCSV = (text: string, source: string): VideoRow[] => {
  const result = Papa.parse<Record<string,string>>(text.replace(/^﻿/,"").trim(), {header:true, skipEmptyLines:true});
  const rows = result.data.filter(r => r.URL && r.Creator);
  const sorted = rows.map(r => {
    const url = (r.URL||"").trim();
    const vid = (url.match(/\/video\/(\d+)/)||[])[1]||"";
    const rev = parseFloat((r["Video Revenue"]||"").replace(/[$,]/g,""))||0;
    const desc = (r.Description||"").replace(/#\w+/g,"").trim();
    const tags = ((r.Description||"").match(/#\w+/g)||[]).join(" ");
    return {
      source,
      creator: (r.Creator||"").trim(),
      videoId: vid,
      videoLink: url,
      revenue: rev,
      itemsSold: parseInt(r["Items Sold"])||0,
      views: parseInt((r["Views Count"]||"").replace(/,/g,""))||0,
      likes: parseInt((r["Likes Count"]||"").replace(/,/g,""))||0,
      comments: parseInt((r["Comments Count"]||"").replace(/,/g,""))||0,
      description: desc,
      hashtags: tags,
      product: (r.Product||"").trim(),
      datePosted: (r.Date||"").slice(0,10),
      audioHook:    C(r.Hooks||""),
      textHook:     "",
      sellingPoints:C(r["Selling Points"]||""),
      visualHook:   "",
      videoLength:  "",
      cta:          "",
      keyIdea:      C(r["Key Idea"] || ""),
      transcript:   C(r.Transcript || ""),
    } as Omit<VideoRow, 'id' | 'rank'>;
  // Include rank in ID so rows with the same videoId (same video under multiple
  // products) each get a unique DB row instead of collapsing via upsert.
  }).sort((a,b)=>b.revenue-a.revenue).map((r,i)=>({...r, id:`${source}_${r.videoId||i}_${i+1}`, rank:i+1}));
  return sorted as VideoRow[];
};

// Apply an overrides map onto an array of raw VideoRows.
const applyOverrides = (rows: VideoRow[], oMap: Map<string, Override>): VideoRow[] =>
  rows.map(r => {
    const ov = oMap.get(r.videoId) || {};
    return {
      ...r,
      audioHook:     ov.audioHook     ?? r.audioHook,
      visualHook:    ov.visualHook    ?? r.visualHook,
      textHook:      ov.textHook      ?? r.textHook,
      videoLength:   ov.videoLength   ?? r.videoLength,
      cta:           ov.cta           ?? r.cta,
      sellingPoints: ov.sellingPoints ?? r.sellingPoints,
      keyIdea:       ov.keyIdea       ?? r.keyIdea,
    };
  });

// ─── EXPORT ───────────────────────────────────────────────────────────────────

const buildXLSX = (at: VideoRow[], lm: VideoRow[], creators: CreatorSummary[]) => {
  const wb = XLSX.utils.book_new();
  const vH = ["#","Creator","Video URL","Revenue ($)","Items Sold","Views","Likes","Comments","Product","Description","Hashtags","Visual Hook","Text Hook","Audio Hook","Video Length","CTA","Selling Points","Transcript"];
  const vR = (r: VideoRow, i: number): (string|number)[] => [i+1,r.creator,r.videoLink,r.revenue,r.itemsSold,r.views,r.likes,r.comments,r.product,r.description,r.hashtags,r.visualHook,r.textHook,r.audioHook,r.videoLength,r.cta,r.sellingPoints,r.transcript||""];
  const mkSheet = (rows: VideoRow[], name: string) => {
    const ws = XLSX.utils.aoa_to_sheet([vH,...rows.map((r,i)=>vR(r,i))]);
    ws["!cols"] = [4,18,30,12,10,12,10,10,28,40,22,22,22,16,12,30,40,60].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb,ws,name);
  };
  mkSheet(at,"All-Time Affiliate"); mkSheet(lm,"Last Month Affiliate");
  const cH=["Creator","Total Videos","Units Sold","GMV ($)","Videos w/ GMV","Videos Last Month","Videos w/ GMV Last Month"];
  const ws3=XLSX.utils.aoa_to_sheet([cH,...creators.map(c=>[c.creator,c.totalVideos,c.unitsSold,c.gmv.toFixed(2),c.videosWithGmv,c.videosLastMonth,c.videosWithGmvLastMonth])]);
  ws3["!cols"]=[22,14,12,16,14,16,22].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb,ws3,"Top Creators");
  XLSX.writeFile(wb,`Ruff-Liners-TikTok-${new Date().toISOString().slice(0,10)}.xlsx`);
};

// ─── TABS ─────────────────────────────────────────────────────────────────────

const TABS = [
  {id:"alltime",   label:"All-Time High GMV",      icon:"🏆"},
  {id:"lastmonth", label:"Last Month High GMV",     icon:"📅"},
  {id:"inhouse",   label:"In-House High GMV",       icon:"🎬"},
  {id:"creators",  label:"Top Creators",            icon:"⭐"},
  {id:"hooks",     label:"Hooks, CTA & Selling Points", icon:"🪝"},
];

const UP_TYPES = [
  {value:"alltime",   label:"All-Time Report  →  feeds All-Time tab + Top Creators"},
  {value:"lastmonth", label:"Last Month Report  →  feeds Last Month tab + monthly creator stats"},
  {value:"inhouse",   label:"In-House Content Report"},
];

// ─── VISUAL HOOK MANAGER MODAL ───────────────────────────────────────────────

function VHManagerModal({ options, onClose, onSave }: {
  options: string[];
  onClose: () => void;
  onSave: (next: string[]) => void;
}) {
  const [items, setItems] = React.useState<string[]>(options);
  const [editIdx, setEditIdx] = React.useState<number|null>(null);
  const [editVal, setEditVal] = React.useState("");
  const [newVal, setNewVal] = React.useState("");

  const startEdit = (i: number) => { setEditIdx(i); setEditVal(items[i]); };
  const commitEdit = (i: number) => {
    const v = editVal.trim();
    if (!v) return;
    setItems(prev => prev.map((x,idx) => idx===i ? v : x));
    setEditIdx(null);
  };
  const deleteItem = (i: number) => {
    setItems(prev => prev.filter((_,idx) => idx!==i));
    if (editIdx===i) setEditIdx(null);
  };
  const addNew = () => {
    const v = newVal.trim();
    if (!v || items.includes(v)) return;
    setItems(prev => [...prev, v]);
    setNewVal("");
  };
  const moveUp = (i: number) => {
    if (i===0) return;
    setItems(prev => { const a=[...prev]; [a[i-1],a[i]]=[a[i],a[i-1]]; return a; });
  };
  const moveDown = (i: number) => {
    setItems(prev => { if (i>=prev.length-1) return prev; const a=[...prev]; [a[i],a[i+1]]=[a[i+1],a[i]]; return a; });
  };

  return (
    <div onClick={e=>{if(e.target===e.currentTarget)onClose();}}
      style={{position:"fixed",inset:0,zIndex:3000,background:"rgba(0,0,0,0.65)",display:"flex",alignItems:"center",justifyContent:"center",padding:"20px"}}>
      <div style={{background:"#fff",borderRadius:14,width:"100%",maxWidth:560,maxHeight:"85vh",display:"flex",flexDirection:"column",boxShadow:"0 20px 60px rgba(0,0,0,0.3)"}}>
        {/* Header */}
        <div style={{padding:"20px 24px 16px",borderBottom:"1px solid #e5e7eb",display:"flex",alignItems:"center",gap:10}}>
          <div style={{fontWeight:800,fontSize:16,color:"#111",flex:1}}>🎬 Manage Visual Hook Options</div>
          <button onClick={onClose} style={{background:"none",border:"none",fontSize:18,cursor:"pointer",color:"#9ca3af",padding:"0 4px",lineHeight:1}}>✕</button>
        </div>

        {/* List */}
        <div style={{flex:1,overflowY:"auto",padding:"12px 24px"}}>
          {items.length===0 && (
            <div style={{textAlign:"center",padding:"32px 0",color:"#9ca3af",fontSize:13}}>No options yet — add one below.</div>
          )}
          {items.map((opt, i) => (
            <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 0",borderBottom:"1px solid #f3f4f6"}}>
              {/* Up/down */}
              <div style={{display:"flex",flexDirection:"column",gap:1}}>
                <button onClick={()=>moveUp(i)} disabled={i===0}
                  style={{background:"none",border:"none",cursor:i===0?"not-allowed":"pointer",color:i===0?"#d1d5db":"#9ca3af",fontSize:10,padding:"1px 3px",lineHeight:1}}>▲</button>
                <button onClick={()=>moveDown(i)} disabled={i===items.length-1}
                  style={{background:"none",border:"none",cursor:i===items.length-1?"not-allowed":"pointer",color:i===items.length-1?"#d1d5db":"#9ca3af",fontSize:10,padding:"1px 3px",lineHeight:1}}>▼</button>
              </div>
              {/* Edit or display */}
              {editIdx===i ? (
                <input autoFocus value={editVal} onChange={e=>setEditVal(e.target.value)}
                  onKeyDown={e=>{ if(e.key==="Enter") commitEdit(i); if(e.key==="Escape") setEditIdx(null); }}
                  style={{flex:1,padding:"5px 9px",border:"1px solid #7c3aed",borderRadius:6,fontFamily:"inherit",fontSize:13,outline:"none"}}/>
              ) : (
                <div style={{flex:1,fontSize:13,color:"#111",wordBreak:"break-word"}}>{opt}</div>
              )}
              {/* Actions */}
              {editIdx===i ? (
                <>
                  <button onClick={()=>commitEdit(i)}
                    style={{padding:"4px 10px",background:"#7c3aed",color:"#fff",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600,whiteSpace:"nowrap"}}>
                    Save
                  </button>
                  <button onClick={()=>setEditIdx(null)}
                    style={{padding:"4px 8px",background:"#f3f4f6",color:"#6b7280",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontSize:12}}>
                    Cancel
                  </button>
                </>
              ) : (
                <>
                  <button onClick={()=>startEdit(i)}
                    style={{padding:"4px 10px",background:"#f3f4f6",color:"#374151",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:500,whiteSpace:"nowrap"}}>
                    ✏️ Edit
                  </button>
                  <button onClick={()=>deleteItem(i)}
                    style={{padding:"4px 8px",background:"none",color:"#dc2626",border:"1px solid #fca5a5",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontSize:12}}>
                    🗑
                  </button>
                </>
              )}
            </div>
          ))}
        </div>

        {/* Add new */}
        <div style={{padding:"14px 24px",borderTop:"1px solid #e5e7eb",display:"flex",gap:8}}>
          <input value={newVal} onChange={e=>setNewVal(e.target.value)}
            onKeyDown={e=>{ if(e.key==="Enter") addNew(); }}
            placeholder="Add new visual hook option…"
            style={{flex:1,padding:"7px 12px",border:"1px solid #d1d5db",borderRadius:7,fontFamily:"inherit",fontSize:13,outline:"none"}}/>
          <button onClick={addNew}
            style={{padding:"7px 16px",background:"#7c3aed",color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontFamily:"inherit",fontSize:13,fontWeight:600,whiteSpace:"nowrap"}}>
            Add
          </button>
        </div>

        {/* Footer */}
        <div style={{padding:"12px 24px",borderTop:"1px solid #e5e7eb",display:"flex",justifyContent:"flex-end",gap:10}}>
          <button onClick={onClose}
            style={{padding:"8px 18px",background:"#f3f4f6",color:"#374151",border:"none",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:13,fontWeight:600}}>
            Cancel
          </button>
          <button onClick={()=>{ onSave(items); onClose(); }}
            style={{padding:"8px 20px",background:"#111",color:"#fff",border:"none",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:13,fontWeight:700}}>
            💾 Save Changes
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── VIDEO CARD ───────────────────────────────────────────────────────────────
// Defined at module level so React never remounts it on parent re-renders.

function VideoCard({ r, showFilter, hiddenIds, editingId, adminMode, transcriptOpen,
  visualHookOptions, toggleHide, cancelEdit, openEdit, saveEdit, toggleTranscript,
  onAddVisualHookOption }: VideoCardProps) {
  const sellPts  = pts(r.sellingPoints).map(lbl);
  const tags     = (r.hashtags||"").split(" ").filter(Boolean);
  const hidden   = hiddenIds.has(r.videoId);
  const isEditing = editingId === r.id;

  const [draft, setDraft] = React.useState<EditDraft>({
    audioHook:"", visualHook:"", textHook:"", videoLength:"", cta:"", sellingPoints:"", keyIdea:""
  });
  const [addingVisualHook, setAddingVisualHook] = React.useState(false);
  const [newVisualHookText, setNewVisualHookText] = React.useState("");
  React.useEffect(() => {
    if (isEditing) {
      setDraft({
        audioHook:    r.audioHook    || "",
        visualHook:   r.visualHook   || "",
        textHook:     r.textHook     || "",
        videoLength:  r.videoLength  || "",
        cta:          r.cta          || "",
        sellingPoints: pts(r.sellingPoints).join("\n"),
        keyIdea:      r.keyIdea      || "",
      });
      setAddingVisualHook(false);
      setNewVisualHookText("");
    }
  }, [isEditing]); // eslint-disable-line react-hooks/exhaustive-deps
  const shortProduct = (r.product||"")
    .replace("for Dogs with Door Protection","")
    .replace("for Full-Size Crew Cab Trucks with Fold Up Seats","")
    .trim();

  const hookRows: [string,string,string][] = [
    ["🎵","Audio Hook", r.audioHook ? hks(r.audioHook).map(lbl).join(" · ") : ""],
    ["🎬","Visual Hook", r.visualHook||""],
    ["🎣","Text Hook",   r.textHook||""],
  ].filter(([,,v]) => v) as [string,string,string][];

  return (
    <div className="rl-card" style={{display:"flex",background:"#fff",borderRadius:14,overflow:"hidden",
      boxShadow:"0 1px 3px rgba(0,0,0,0.07)",
      border:`1px solid ${hidden&&adminMode?"#fca5a5":"#e5e7eb"}`,
      marginBottom:16, opacity:hidden&&adminMode?0.55:1}}>

      <div className="rl-card-vid" style={{flexShrink:0,width:325,background:"#0a0a0a"}}>
        {r.videoId ? (
          <div className="rl-embed" style={{width:325,height:578,overflow:"hidden",flexShrink:0}}>
            <iframe className="rl-iframe" src={`https://www.tiktok.com/embed/v2/${r.videoId}`}
              style={{display:"block",width:325,height:738,border:"none"}}
              allowFullScreen allow="encrypted-media" loading="lazy" title={`@${r.creator}`}/>
          </div>
        ) : (
          <div className="rl-embed" style={{width:325,height:578,display:"flex",alignItems:"center",justifyContent:"center",color:"#444",flexDirection:"column",gap:6,fontSize:12}}>
            <span style={{fontSize:24}}>📹</span>No embed
          </div>
        )}
      </div>

      <div className="rl-card-body" style={{flex:1,padding:"16px 20px",overflow:"hidden",minWidth:0,display:"flex",flexDirection:"column",gap:10}}>

        <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:10}}>
          <div style={{display:"flex",alignItems:"center",gap:9,minWidth:0}}>
            <div style={{background:"#111",color:"#fff",borderRadius:7,minWidth:30,height:30,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,flexShrink:0}}>#{r.rank}</div>
            <div style={{minWidth:0}}>
              <div style={{fontWeight:700,fontSize:14,color:"#111"}}>@{r.creator}</div>
              <div style={{fontSize:11,color:"#9ca3af",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{shortProduct} · {r.datePosted}</div>
            </div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
            <div style={{textAlign:"right"}}>
              <div style={{fontWeight:800,fontSize:22,color:"#16a34a",lineHeight:1}}>{f$(r.revenue)}</div>
              <div style={{fontSize:11,color:"#9ca3af"}}>{fN(r.itemsSold)} sold</div>
            </div>
            {showFilter && (
              <button onClick={()=>toggleHide(r.videoId)}
                style={{padding:"4px 9px",borderRadius:7,border:"1px solid",borderColor:hidden?"#fca5a5":"#d1d5db",background:hidden?"#fff5f5":"#fff",cursor:"pointer",fontFamily:"inherit",fontSize:11,color:hidden?"#dc2626":"#6b7280"}}>
                {hidden?"🚫":"👁"}
              </button>
            )}
            {adminMode && (
              <button onClick={()=>isEditing?cancelEdit():openEdit(r)}
                style={{padding:"5px 12px",borderRadius:7,border:"1px solid",borderColor:isEditing?"#3b82f6":"#d1d5db",background:isEditing?"#eff6ff":"#fff",cursor:"pointer",fontFamily:"inherit",fontSize:11,fontWeight:600,color:isEditing?"#2563eb":"#374151"}}>
                {isEditing?"✕ Close":"✏️ Edit"}
              </button>
            )}
          </div>
        </div>

        {isEditing && (
          <div style={{background:"#f8faff",border:"1px solid #bfdbfe",borderRadius:10,padding:"14px 16px",display:"flex",flexDirection:"column",gap:10}}>
            <div style={{fontSize:10,fontWeight:700,color:"#2563eb",textTransform:"uppercase",letterSpacing:"0.07em"}}>Edit Fields</div>
            {(
              [
                ["audioHook",   "🎵 Audio Hook",   "textarea"],
                ["textHook",    "🎣 Text Hook",    "textarea"],
                ["videoLength", "⏱ Video Length",  "input"  ],
                ["cta",         "📣 Call to Action","input"  ],
                ["keyIdea",     "💡 Key Idea",      "textarea"],
              ] as [keyof EditDraft, string, string][]
            ).map(([field, label, type]) => (
              <div key={field}>
                <div style={{fontSize:10,color:"#6b7280",fontWeight:600,marginBottom:3}}>{label}</div>
                {type==="input" ? (
                  <input value={draft[field]||""} onChange={e=>setDraft(d=>({...d,[field]:e.target.value}))}
                    style={{width:"100%",padding:"6px 10px",border:"1px solid #d1d5db",borderRadius:6,fontFamily:"inherit",fontSize:12,boxSizing:"border-box",outline:"none"}}/>
                ) : (
                  <textarea value={draft[field]||""} onChange={e=>setDraft(d=>({...d,[field]:e.target.value}))} rows={2}
                    style={{display:"block",width:"100%",padding:"6px 10px",border:"1px solid #d1d5db",borderRadius:6,fontFamily:"inherit",fontSize:12,resize:"vertical",boxSizing:"border-box",outline:"none",lineHeight:1.5}}/>
                )}
              </div>
            ))}
            {/* ── Visual Hook dropdown ── */}
            <div>
              <div style={{fontSize:10,color:"#6b7280",fontWeight:600,marginBottom:3}}>🎬 Visual Hook</div>
              {!addingVisualHook ? (
                <div style={{display:"flex",gap:6}}>
                  <select
                    value={draft.visualHook||""}
                    onChange={e=>{
                      if (e.target.value==="__add_new__") { setAddingVisualHook(true); }
                      else { setDraft(d=>({...d,visualHook:e.target.value})); }
                    }}
                    style={{flex:1,padding:"6px 10px",border:"1px solid #d1d5db",borderRadius:6,fontFamily:"inherit",fontSize:12,background:"#fff",color:"#111",outline:"none",cursor:"pointer"}}>
                    <option value="">— none —</option>
                    {[...visualHookOptions].sort((a,b)=>a.localeCompare(b)).map(opt=>(
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                    <option value="__add_new__">+ Add new option…</option>
                  </select>
                </div>
              ) : (
                <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                  <input
                    autoFocus
                    value={newVisualHookText}
                    onChange={e=>setNewVisualHookText(e.target.value)}
                    placeholder="Type new visual hook…"
                    onKeyDown={e=>{
                      if (e.key==="Enter" && newVisualHookText.trim()) {
                        const v=newVisualHookText.trim();
                        onAddVisualHookOption(v);
                        setDraft(d=>({...d,visualHook:v}));
                        setNewVisualHookText(""); setAddingVisualHook(false);
                      }
                      if (e.key==="Escape") { setAddingVisualHook(false); setNewVisualHookText(""); }
                    }}
                    style={{flex:1,padding:"6px 10px",border:"1px solid #7c3aed",borderRadius:6,fontFamily:"inherit",fontSize:12,outline:"none",minWidth:0}}/>
                  <button
                    type="button"
                    onClick={()=>{
                      const v=newVisualHookText.trim();
                      if (!v) return;
                      onAddVisualHookOption(v);
                      setDraft(d=>({...d,visualHook:v}));
                      setNewVisualHookText(""); setAddingVisualHook(false);
                    }}
                    style={{padding:"6px 12px",background:"#7c3aed",color:"#fff",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600,whiteSpace:"nowrap"}}>
                    Add
                  </button>
                  <button
                    type="button"
                    onClick={()=>{ setAddingVisualHook(false); setNewVisualHookText(""); }}
                    style={{padding:"6px 10px",background:"#f3f4f6",color:"#6b7280",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontSize:12}}>
                    Cancel
                  </button>
                </div>
              )}
            </div>
            <div>
              <div style={{fontSize:10,color:"#6b7280",fontWeight:600,marginBottom:3}}>✅ Selling Points <span style={{fontWeight:400,color:"#9ca3af"}}>— one per line</span></div>
              <textarea value={draft.sellingPoints||""} onChange={e=>setDraft(d=>({...d,sellingPoints:e.target.value}))} rows={3}
                style={{display:"block",width:"100%",padding:"6px 10px",border:"1px solid #d1d5db",borderRadius:6,fontFamily:"inherit",fontSize:12,resize:"vertical",boxSizing:"border-box",outline:"none",lineHeight:1.6}}/>
            </div>
            <button onClick={()=>saveEdit(r, draft)}
              style={{alignSelf:"flex-start",padding:"7px 20px",background:"#111",color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:700}}>
              💾 Save Changes
            </button>
          </div>
        )}

        <div style={{display:"flex",gap:18,padding:"8px 0",borderTop:"1px solid #f3f4f6",borderBottom:"1px solid #f3f4f6"}}>
          {[["👁","Views",fK(r.views)],["❤️","Likes",fK(r.likes)],["💬","Comments",fK(r.comments)]].map(([ic,lb,v])=>(
            <div key={lb as string}>
              <div style={{fontSize:9,color:"#9ca3af",textTransform:"uppercase",letterSpacing:"0.04em",marginBottom:1}}>{ic} {lb}</div>
              <div style={{fontWeight:700,fontSize:16,color:"#111"}}>{v}</div>
            </div>
          ))}
        </div>

        {r.product && (
          <div style={{display:"flex",alignItems:"center",gap:6}}>
            <span style={{fontSize:10,fontWeight:700,color:"#374151",textTransform:"uppercase",letterSpacing:"0.05em",flexShrink:0}}>🛍</span>
            <span style={{fontSize:11,color:"#fff",background:"#374151",borderRadius:20,padding:"2px 10px",fontWeight:500}}>{r.product}</span>
          </div>
        )}

        {(r.description || tags.length > 0) && (
          <div>
            <div style={{fontSize:10,fontWeight:700,color:"#374151",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:4}}>📝 Description</div>
            <div style={{fontSize:12,lineHeight:1.6}}>
              {r.description && <span style={{color:"#4b5563",fontStyle:"italic"}}>"{ r.description}" </span>}
              {tags.map(t=><span key={t} style={{color:"#2563eb",fontWeight:500,marginRight:4}}>{t}</span>)}
            </div>
          </div>
        )}

        {hookRows.length>0 && (
          <div>
            <div style={{fontSize:10,fontWeight:700,color:"#374151",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:5}}>Hooks</div>
            <div style={{display:"flex",flexDirection:"column",gap:4}}>
              {hookRows.map(([ic,lb,val])=>(
                <div key={lb} style={{display:"flex",gap:7,alignItems:"flex-start",fontSize:12}}>
                  <span style={{flexShrink:0,width:14,textAlign:"center"}}>{ic}</span>
                  <span style={{color:"#6b7280",flexShrink:0,fontWeight:600}}>{lb}:</span>
                  <span style={{color:"#111",lineHeight:1.45}}>{val}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        {r.keyIdea && (
          <div>
            <div style={{fontSize:10,fontWeight:700,color:"#374151",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:4}}>💡 Key Idea</div>
            <div style={{fontSize:12,color:"#374151",lineHeight:1.5,background:"#fefce8",border:"1px solid #fef08a",borderRadius:7,padding:"7px 11px"}}>{r.keyIdea}</div>
          </div>
        )}

        {r.cta && (
          <div>
            <div style={{fontSize:10,fontWeight:700,color:"#374151",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:4}}>📣 Call to Action</div>
            <div style={{fontSize:12,color:"#374151",lineHeight:1.5,background:"#fdf4ff",border:"1px solid #e9d5ff",borderRadius:7,padding:"7px 11px"}}>{r.cta}</div>
          </div>
        )}

        {sellPts.length>0 && (
          <div>
            <div style={{fontSize:10,fontWeight:700,color:"#374151",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:5}}>✅ Selling Points</div>
            <div style={{display:"flex",flexDirection:"column",gap:3}}>
              {sellPts.map((p,i)=>(
                <div key={i} style={{display:"flex",gap:7,alignItems:"flex-start"}}>
                  <span style={{color:"#16a34a",flexShrink:0,fontWeight:700}}>✓</span>
                  <span style={{fontSize:12,color:"#374151",lineHeight:1.45}}>{p}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        {r.transcript && (
          <div>
            <button onClick={()=>toggleTranscript(r.id)}
              style={{display:"inline-flex",alignItems:"center",gap:5,fontSize:11,color:"#374151",background:"none",border:"1px solid #d1d5db",borderRadius:6,padding:"4px 11px",cursor:"pointer",fontFamily:"inherit",fontWeight:500}}>
              📄 {transcriptOpen.has(r.id)?"Hide Transcript":"View Transcript"}
            </button>
            {transcriptOpen.has(r.id) && (
              <div style={{marginTop:8,background:"#f9fafb",border:"1px solid #e5e7eb",borderRadius:8,padding:"12px 14px",maxHeight:260,overflowY:"auto"}}>
                <div style={{fontSize:10,fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:6}}>Transcript</div>
                <div style={{fontSize:12,color:"#374151",lineHeight:1.75,whiteSpace:"pre-wrap"}}>{r.transcript}</div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── COMPONENT ────────────────────────────────────────────────────────────────

// Upsert rows in chunks to avoid PostgREST payload limits. Returns true on
// full success, false (with console error) if any chunk fails.
async function upsertReports(rows: Record<string, unknown>[]): Promise<boolean> {
  const CHUNK = 100;
  for (let i = 0; i < rows.length; i += CHUNK) {
    const { error } = await supabase.from('tiktok_reports').upsert(rows.slice(i, i + CHUNK), { onConflict: 'id' });
    if (error) { console.error('upsertReports chunk error:', error); return false; }
  }
  return true;
}

export default function TikTokShopReporter() {

  const [tab,       setTab]       = useState("alltime");
  const [allTime,   setAllTime]   = useState<VideoRow[]>([]);
  const [lastMonth, setLastMonth] = useState<VideoRow[]>([]);
  const [inhouse,   setInhouse]   = useState<VideoRow[]>([]);
  const [adminMode, setAdminMode] = useState(false);
  const [isAdmin,   setIsAdmin]   = useState(false);
  const [hiddenIds, setHiddenIds] = useState<Set<string>>(new Set());
  const [editingId,  setEditingId]  = useState<string | null>(null);
  const [transcriptOpen, setTranscriptOpen] = useState<Set<string>>(new Set());
  const [upType,    setUpType]    = useState("alltime");
  const [drag,      setDrag]      = useState(false);
  const [showUp,    setShowUp]    = useState(false);
  const [pageAt,    setPageAt]    = useState(20);
  const [pageLm,    setPageLm]    = useState(20);
  const [savedAt,   setSavedAt]   = useState(0);
  const [savedLm,   setSavedLm]   = useState(0);
  const [savedCr,   setSavedCr]   = useState(0);
  const [draftAt,   setDraftAt]   = useState<number|"">("");
  const [draftLm,   setDraftLm]   = useState<number|"">("");
  const [draftCr,   setDraftCr]   = useState<number|"">("");
  const [overridesMap, setOverridesMap] = useState<Map<string, Override>>(new Map());
  const localOverridesRef = useRef<Map<string, Override>>(new Map());
  const [dataLoading, setDataLoading] = useState(true);
  const fileRef = useRef<HTMLInputElement>(null);
  const xlsxImportRef = useRef<HTMLInputElement>(null);
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [showLoginModal, setShowLoginModal] = useState(false);
  const [lastImported, setLastImported] = useState("");
  const [loginEmail, setLoginEmail] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [loginLoading, setLoginLoading] = useState(false);
  const [loginError, setLoginError] = useState('');
  const [pubAt, setPubAt] = useState(0);
  const [pubLm, setPubLm] = useState(0);
  const [pubCr, setPubCr] = useState(0);
  const [pubHiddenIds, setPubHiddenIds] = useState<Set<string>>(new Set());
  const [pubAllTime,   setPubAllTime]   = useState<VideoRow[]>([]);
  const [pubLastMonth, setPubLastMonth] = useState<VideoRow[]>([]);
  const [pubInhouse,   setPubInhouse]   = useState<VideoRow[]>([]);
  const [publishing, setPublishing] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [uploadStep, setUploadStep] = useState("");
  const [atAgg, setAtAgg] = useState<{
    creators: CreatorSummary[];
    visualHooks: HookSummary[];
    textHooks: HookSummary[];
    audioHooks: HookSummary[];
    ctas: HookSummary[];
    sellingPoints: SellingPointSummary[];
  } | null>(null);
  const [pubAtAgg, setPubAtAgg] = useState<typeof atAgg>(null);
  const [visualHookOptions, setVisualHookOptions] = useState<string[]>([]);
  const [showVHManager, setShowVHManager] = useState(false);

  // ── initial load from Supabase ──────────────────────────────────────────────

  useEffect(() => {
    const toRow = (r: Record<string,unknown>, om: Map<string,Override>): VideoRow => {
      const ov = om.get(r.video_id as string) || {};
      return {
        id:           r.id as string,
        source:       r.source as string,
        videoId:      r.video_id as string,
        videoLink:    r.video_link as string,
        creator:      r.creator as string,
        revenue:      r.revenue as number,
        itemsSold:    r.items_sold as number,
        views:        r.views as number,
        likes:        r.likes as number,
        comments:     r.comments as number,
        description:  r.description as string,
        hashtags:     r.hashtags as string,
        product:      r.product as string,
        datePosted:   r.date_posted as string,
        audioHook:    ov.audioHook    ?? r.audio_hook as string ?? "",
        visualHook:   ov.visualHook   ?? "",
        textHook:     ov.textHook     ?? "",
        videoLength:  ov.videoLength  ?? "",
        cta:          ov.cta          ?? "",
        sellingPoints:ov.sellingPoints ?? r.selling_points as string ?? "",
        keyIdea:      ov.keyIdea      ?? r.key_idea as string ?? "",
        transcript:   r.transcript as string ?? "",
        rank:         r.rank as number,
      };
    };
    const srt = (a: Record<string,unknown>, b: Record<string,unknown>) => (b.revenue as number) - (a.revenue as number);

    const CACHE_KEY = 'rl_data_v1';

    // Restore from sessionStorage immediately — skips the loading spinner on tab switches / page reloads
    try {
      const raw = sessionStorage.getItem(CACHE_KEY);
      if (raw) {
        const c = JSON.parse(raw);
        setAllTime(c.at||[]); setLastMonth(c.lm||[]); setInhouse(c.inh||[]);
        setPubAllTime(c.pat||[]); setPubLastMonth(c.plm||[]); setPubInhouse(c.pinh||[]);
        setHiddenIds(new Set(c.hid||[]));
        setPubHiddenIds(new Set(c.phid||[]));
        setOverridesMap(new Map(c.omap||[]));
        setSavedAt(c.sat||0); setDraftAt(c.sat||"");
        setSavedLm(c.slm||0); setDraftLm(c.slm||"");
        setSavedCr(c.scr||0); setDraftCr(c.scr||"");
        setPubAt(c.npat||0); setPubLm(c.nplm||0); setPubCr(c.npcr||0);
        setLastImported(c.li||"");
        if (c.atAgg) setAtAgg(c.atAgg);
        if (c.pubAtAgg) setPubAtAgg(c.pubAtAgg);
        setDataLoading(false);
      }
    } catch {}

    const load = async () => {
      const { data: { session } } = await supabase.auth.getSession();
      const loggedIn = !!session;
      setIsLoggedIn(loggedIn);
      const adminFlag = session?.user?.user_metadata?.is_admin === true;
      setIsAdmin(adminFlag);

      // Run all reads in parallel, each source in its own query to avoid the 1000-row default limit
      const [
        { data: overrideRows },
        { data: hiddenRows },
        { data: settings },
        { data: atRows },
        { data: lmRows },
        { data: inhRows },
        { data: pubAtRows },
        { data: pubLmRows },
        { data: pubInhRows },
      ] = await Promise.all([
        supabase.from('tiktok_overrides').select('*').order('updated_at', { ascending: true }),
        supabase.from('tiktok_hidden_videos').select('video_id'),
        supabase.from('tiktok_hub_settings').select('*'),
        supabase.from('tiktok_reports').select('*').eq('source','alltime').limit(10000),
        supabase.from('tiktok_reports').select('*').eq('source','lastmonth').limit(10000),
        supabase.from('tiktok_reports').select('*').eq('source','inhouse').limit(10000),
        supabase.from('tiktok_reports').select('*').eq('source','pub_alltime').limit(10000),
        supabase.from('tiktok_reports').select('*').eq('source','pub_lastmonth').limit(10000),
        supabase.from('tiktok_reports').select('*').eq('source','pub_inhouse').limit(10000),
      ]);

      const newMap = new Map<string, Override>();
      overrideRows?.forEach((o: Record<string,string>) => {
        newMap.set(o.report_id, {
          audioHook:    o.audio_hook    || undefined,
          visualHook:   o.visual_hook   || undefined,
          textHook:     o.text_hook     || undefined,
          videoLength:  o.video_length  || undefined,
          cta:          o.cta           || undefined,
          sellingPoints:o.selling_points|| undefined,
          keyIdea:      o.key_idea      || undefined,
        });
      });
      // Merge locally-saved overrides so a background load() never reverts unsaved edits
      localOverridesRef.current.forEach((v, k) => newMap.set(k, v));
      setOverridesMap(newMap);

      const hiddenSet = new Set(hiddenRows?.map((h: {video_id:string}) => h.video_id) || []);
      setHiddenIds(hiddenSet);

      let sat=0, slm=0, scr=0, npat=0, nplm=0, npcr=0, phid: string[]=[], li="";
      settings?.forEach((s: {key:string, value:string}) => {
        const n = Number(s.value) || 0;
        if (s.key === 'filter_at') { sat=n; setSavedAt(n); setDraftAt(n || ""); }
        if (s.key === 'filter_lm') { slm=n; setSavedLm(n); setDraftLm(n || ""); }
        if (s.key === 'filter_cr') { scr=n; setSavedCr(n); setDraftCr(n || ""); }
        if (s.key === 'pub_filter_at') { npat=n; setPubAt(n); }
        if (s.key === 'pub_filter_lm') { nplm=n; setPubLm(n); }
        if (s.key === 'pub_filter_cr') { npcr=n; setPubCr(n); }
        if (s.key === 'pub_hidden_ids') { try { phid=JSON.parse(s.value); setPubHiddenIds(new Set(phid)); } catch {} }
        if (s.key === 'last_imported') { li=s.value; setLastImported(s.value); }
        // ov_VIDEOID keys are override backups written by saveEdit
        if (s.key.startsWith('ov_')) {
          try {
            const videoId = s.key.slice(3);
            const f: Override = JSON.parse(s.value);
            newMap.set(videoId, { ...(newMap.get(videoId) || {}), ...f });
          } catch {}
        }
      });

      const atAggSetting = settings?.find((s: {key:string,value:string}) => s.key === 'at_agg');
      if (atAggSetting) { try { setAtAgg(JSON.parse(atAggSetting.value)); } catch {} }
      const pubAtAggSetting = settings?.find((s: {key:string,value:string}) => s.key === 'pub_at_agg');
      if (pubAtAggSetting) { try { setPubAtAgg(JSON.parse(pubAtAggSetting.value)); } catch {} }

      const at   = (atRows    || []).sort(srt).map(r => toRow(r, newMap));
      const lm   = (lmRows    || []).sort(srt).map(r => toRow(r, newMap));
      const inh  = (inhRows   || []).sort(srt).map(r => toRow(r, newMap));
      const pat  = (pubAtRows || []).sort(srt).map(r => toRow(r, newMap));
      const plm  = (pubLmRows || []).sort(srt).map(r => toRow(r, newMap));
      const pinh = (pubInhRows|| []).sort(srt).map(r => toRow(r, newMap));

      // ── Visual hook options: load from settings or seed from top 17 ──
      const vhoSetting = settings?.find((s: {key:string,value:string}) => s.key === 'visual_hook_options');
      if (vhoSetting) {
        try { setVisualHookOptions(JSON.parse(vhoSetting.value)); } catch {}
      } else if (adminFlag && at.length >= 1) {
        // First time: seed options from top 17 videos' visual hooks (non-empty, deduplicated)
        const seeded: string[] = [];
        at.slice(0, 17).forEach(row => {
          const v = (row.visualHook || "").trim();
          if (v && !seeded.includes(v)) seeded.push(v);
        });
        setVisualHookOptions(seeded);
        supabase.from('tiktok_hub_settings').upsert({
          key: 'visual_hook_options',
          value: JSON.stringify(seeded),
          updated_at: new Date().toISOString(),
        });
        // Clear visual hooks for videos ranked 18+ (rank is 1-based from CSV, but
        // array position after sort is 0-based — clear index 17 and beyond)
        if (at.length > 17) {
          const toClear = at.slice(17).filter(row => row.visualHook);
          if (toClear.length > 0) {
            const clearOps = toClear.map(row => ({
              report_id: row.videoId,
              visual_hook: null,
              updated_at: new Date().toISOString(),
            }));
            supabase.from('tiktok_overrides').upsert(clearOps, { onConflict: 'report_id' });
          }
        }
      }

      setAllTime(at); setLastMonth(lm); setInhouse(inh);
      setPubAllTime(pat); setPubLastMonth(plm); setPubInhouse(pinh);
      setDataLoading(false);

      // Persist to sessionStorage so next load is instant
      try {
        const atAggVal = atAggSetting ? (() => { try { return JSON.parse(atAggSetting.value); } catch { return null; } })() : null;
        const pubAtAggVal = pubAtAggSetting ? (() => { try { return JSON.parse(pubAtAggSetting.value); } catch { return null; } })() : null;
        sessionStorage.setItem(CACHE_KEY, JSON.stringify({
          at, lm, inh, pat, plm, pinh,
          hid: Array.from(hiddenSet),
          phid,
          omap: Array.from(newMap.entries()),
          sat, slm, scr, npat, nplm, npcr, li,
          atAgg: atAggVal, pubAtAgg: pubAtAggVal,
        }));
      } catch {}
    };

    // Initial load — show spinner only if no cached data was restored above
    load();

    // Re-load when auth state changes (silent refresh in background)
    const { data: { subscription: authSub } } = supabase.auth.onAuthStateChange((event, session) => {
      const loggedIn = !!session;
      setIsLoggedIn(loggedIn);
      setIsAdmin(session?.user?.user_metadata?.is_admin === true);
      if (!loggedIn) { setAdminMode(false); setShowLoginModal(false); }
      if (event === 'SIGNED_IN' || event === 'TOKEN_REFRESHED') load();
    });

    return () => authSub.unsubscribe();
  }, []);

// ── derived ─────────────────────────────────────────────────────────────────

  const visAllTime = useMemo(
    () => {
      if (adminMode) {
        return savedAt > 0 ? allTime.filter(v => v.revenue >= savedAt) : allTime;
      }
      const base = pubAllTime.filter(v => !pubHiddenIds.has(v.videoId));
      return pubAt > 0 ? base.filter(v => v.revenue >= pubAt) : base;
    },
    [allTime, adminMode, savedAt, pubAllTime, pubHiddenIds, pubAt]
  );

  const buildCreators = (atV: VideoRow[], lmV: VideoRow[], inhV: VideoRow[]): CreatorSummary[] => {
    if (!atV.length && !lmV.length && !inhV.length) return [];
    const map: Record<string, {creator:string; allV:VideoRow[]; lmV:VideoRow[]; inhV:VideoRow[]}> = {};
    const addTo = (arr: VideoRow[], key: 'allV'|'lmV'|'inhV') => arr.forEach(v => {
      const k = v.creator.toLowerCase().trim();
      if (!map[k]) map[k] = {creator:v.creator, allV:[], lmV:[], inhV:[]};
      map[k][key].push(v);
    });
    addTo(atV, 'allV'); addTo(lmV, 'lmV'); addTo(inhV, 'inhV');
    return Object.values(map)
      .map(c => {
        const top3 = [...c.allV].sort((a,b)=>b.revenue-a.revenue).slice(0,3);
        // Unique total videos across all three sources (deduplicated by videoId)
        const seen = new Set<string>();
        [...c.allV, ...c.lmV, ...c.inhV].forEach(v => seen.add(v.videoId || v.id));
        return {
          creator:               c.creator,
          totalVideos:           seen.size,
          unitsSold:             c.allV.reduce((s,v)=>s+v.itemsSold,0),
          gmv:                   c.allV.reduce((s,v)=>s+v.revenue,0),
          videosWithGmv:         c.allV.filter(v=>v.revenue>0).length,
          videosLastMonth:       c.lmV.length,
          videosWithGmvLastMonth:c.lmV.filter(v=>v.revenue>0).length,
          top3,
        };
      })
      .filter(c => c.gmv > 0)
      .sort((a,b) => b.gmv - a.gmv);
  };
  const creators    = useMemo(() => buildCreators(allTime,    lastMonth,    inhouse),    [allTime,    lastMonth,    inhouse]);    // eslint-disable-line react-hooks/exhaustive-deps
  const pubCreators = useMemo(() => buildCreators(pubAllTime, pubLastMonth, pubInhouse), [pubAllTime, pubLastMonth, pubInhouse]); // eslint-disable-line react-hooks/exhaustive-deps

  const filteredLastMonth = useMemo(
    () => {
      if (adminMode) return savedLm > 0 ? lastMonth.filter(v => v.revenue >= savedLm) : lastMonth;
      const base = pubLastMonth.filter(v => !pubHiddenIds.has(v.videoId));
      return pubLm > 0 ? base.filter(v => v.revenue >= pubLm) : base;
    },
    [lastMonth, adminMode, savedLm, pubLastMonth, pubLm, pubHiddenIds]
  );

  const visInhouse = useMemo(
    () => adminMode ? inhouse : pubInhouse.filter(v => !pubHiddenIds.has(v.videoId)),
    [inhouse, adminMode, pubInhouse, pubHiddenIds]
  );

  const filteredCreators = useMemo(() => {
    // In creator mode fall back to atAgg when pubAtAgg hasn't been published yet
    const aggSrc = adminMode ? atAgg?.creators : (pubAtAgg?.creators ?? atAgg?.creators);
    const rowSrc = adminMode ? creators : pubCreators;
    const src = aggSrc || rowSrc;
    // If using agg creators, update videosLastMonth from current lastMonth state
    const lmRows = adminMode ? lastMonth : pubLastMonth;
    const enhanced = aggSrc ? src.map(c => {
      const key = c.creator.toLowerCase().trim();
      const lmVids = lmRows.filter(v => v.creator.toLowerCase().trim() === key);
      return { ...c, videosLastMonth: lmVids.length, videosWithGmvLastMonth: lmVids.filter(v => v.revenue > 0).length };
    }) : src;
    const threshold = adminMode ? savedCr : pubCr;
    return threshold > 0 ? enhanced.filter(c => c.gmv >= threshold) : enhanced;
  }, [atAgg, pubAtAgg, creators, pubCreators, lastMonth, pubLastMonth, adminMode, savedCr, pubCr]);

  const buildTopHooks = (videos: VideoRow[], field: 'visualHook' | 'textHook' | 'cta'): HookSummary[] => {
    const map: Record<string, { hookText: string; videos: VideoRow[] }> = {};
    videos.forEach(v => {
      const raw = (v[field] || "").trim();
      if (!raw) return;
      const key = raw.toLowerCase();
      if (!map[key]) map[key] = { hookText: raw, videos: [] };
      map[key].videos.push(v);
    });
    return Object.values(map)
      .map(h => {
        const topVideos    = [...h.videos].sort((a, b) => b.revenue - a.revenue).slice(0, 3);
        const totalGmv     = h.videos.reduce((s, v) => s + v.revenue, 0);
        const totalViews   = h.videos.reduce((s, v) => s + v.views, 0);
        const totalUnitsSold = h.videos.reduce((s, v) => s + v.itemsSold, 0);
        const videosWithGmv  = h.videos.filter(v => v.revenue > 0).length;
        return { hookText: h.hookText, totalVideos: h.videos.length, totalViews, totalUnitsSold, totalGmv, videosWithGmv, topVideos };
      })
      .filter(h => h.totalGmv > 0)
      .sort((a, b) => b.totalGmv - a.totalGmv)
      .slice(0, 3);
  };

  const buildTopAudioHooks = (videos: VideoRow[]): HookSummary[] => {
    const map: Record<string, { hookText: string; videos: VideoRow[] }> = {};
    videos.forEach(v => {
      const hooks = pts(v.audioHook).map(lbl).filter(Boolean);
      hooks.forEach(h => {
        const key = h.toLowerCase().trim();
        if (!key) return;
        if (!map[key]) map[key] = { hookText: h, videos: [] };
        if (!map[key].videos.includes(v)) map[key].videos.push(v);
      });
    });
    return Object.values(map)
      .map(h => {
        const topVideos      = [...h.videos].sort((a, b) => b.revenue - a.revenue).slice(0, 3);
        const totalGmv       = h.videos.reduce((s, v) => s + v.revenue, 0);
        const totalViews     = h.videos.reduce((s, v) => s + v.views, 0);
        const totalUnitsSold = h.videos.reduce((s, v) => s + v.itemsSold, 0);
        const videosWithGmv  = h.videos.filter(v => v.revenue > 0).length;
        return { hookText: h.hookText, totalVideos: h.videos.length, totalViews, totalUnitsSold, totalGmv, videosWithGmv, topVideos };
      })
      .filter(h => h.totalGmv > 0)
      .sort((a, b) => b.totalGmv - a.totalGmv)
      .slice(0, 3);
  };

  const buildTopSellingPoints = (videos: VideoRow[]): SellingPointSummary[] => {
    const map: Record<string, { point: string; product: string; videos: VideoRow[] }> = {};
    videos.forEach(v => {
      const product = (v.product || "").trim();
      pts(v.sellingPoints).forEach(raw => {
        const cleaned = lbl(raw).trim();
        if (!cleaned) return;
        const key = `${product}::${cleaned.toLowerCase()}`;
        if (!map[key]) map[key] = { point: cleaned, product, videos: [] };
        if (!map[key].videos.includes(v)) map[key].videos.push(v);
      });
    });
    return Object.values(map)
      .map(h => {
        const topVideos      = [...h.videos].sort((a, b) => b.revenue - a.revenue).slice(0, 3);
        const totalGmv       = h.videos.reduce((s, v) => s + v.revenue, 0);
        const totalViews     = h.videos.reduce((s, v) => s + v.views, 0);
        const totalUnitsSold = h.videos.reduce((s, v) => s + v.itemsSold, 0);
        const videosWithGmv  = h.videos.filter(v => v.revenue > 0).length;
        return { point: h.point, product: h.product, totalVideos: h.videos.length, totalViews, totalUnitsSold, totalGmv, videosWithGmv, topVideos };
      })
      .filter(h => h.totalGmv > 0)
      .sort((a, b) => b.totalGmv - a.totalGmv)
      .slice(0, 5);
  };

  // Always compute live from the current rows so manual edits (overrides) are immediately
  // reflected without waiting for a CSV re-upload or atAgg recompute.
  const src = adminMode ? allTime : pubAllTime;
  const topVisualHooks   = useMemo(() => buildTopHooks(src, 'visualHook'),  [src]); // eslint-disable-line react-hooks/exhaustive-deps
  const topTextHooks     = useMemo(() => buildTopHooks(src, 'textHook'),    [src]); // eslint-disable-line react-hooks/exhaustive-deps
  const topAudioHooks    = useMemo(() => buildTopAudioHooks(src),           [src]); // eslint-disable-line react-hooks/exhaustive-deps
  const topCTAs          = useMemo(() => buildTopHooks(src, 'cta'),         [src]); // eslint-disable-line react-hooks/exhaustive-deps
  const topSellingPoints = useMemo(() => buildTopSellingPoints(src),        [src]); // eslint-disable-line react-hooks/exhaustive-deps

  // ── filter helpers ───────────────────────────────────────────────────────────

  const persistSetting = async (key: string, n: number) => {
    if (n > 0) {
      await supabase.from('tiktok_hub_settings').upsert({ key, value: String(n), updated_at: new Date().toISOString() });
    } else {
      await supabase.from('tiktok_hub_settings').delete().eq('key', key);
    }
  };

  const saveFilterAt = async () => { const n=Math.max(0,Number(draftAt)||0); setSavedAt(n); setDraftAt(n||""); setPageAt(20); await persistSetting('filter_at',n); };
  const saveFilterLm = async () => { const n=Math.max(0,Number(draftLm)||0); setSavedLm(n); setDraftLm(n||""); setPageLm(20); await persistSetting('filter_lm',n); };
  const saveFilterCr = async () => { const n=Math.max(0,Number(draftCr)||0); setSavedCr(n); setDraftCr(n||""); await persistSetting('filter_cr',n); };
  const clearFilterAt = () => { setSavedAt(0); setDraftAt(""); setPageAt(20); persistSetting('filter_at',0); };
  const clearFilterLm = () => { setSavedLm(0); setDraftLm(""); setPageLm(20); persistSetting('filter_lm',0); };
  const clearFilterCr = () => { setSavedCr(0); setDraftCr(""); persistSetting('filter_cr',0); };

  const publishDashboard = async () => {
    setPublishing(true);
    const now = new Date().toISOString();

    // Use position-based pub IDs so rows with duplicate videoIds never conflict
    // within a chunk regardless of which ID format the source rows use.
    const toDbRow = (r: VideoRow, pubSource: string, idx: number) => ({
      id: `${pubSource}_${r.videoId||"x"}_${idx+1}`,
      source: pubSource,
      video_id:      r.videoId,
      video_link:    r.videoLink,
      creator:       r.creator,
      revenue:       r.revenue,
      items_sold:    r.itemsSold,
      views:         r.views,
      likes:         r.likes,
      comments:      r.comments,
      description:   r.description,
      hashtags:      r.hashtags,
      product:       r.product,
      date_posted:   r.datePosted,
      audio_hook:    r.audioHook,
      selling_points:r.sellingPoints,
      key_idea:      r.keyIdea,
      transcript:    r.transcript,
      rank:          r.rank,
    });

    // Update each published source independently so uploading one CSV never
    // wipes another source that wasn't re-uploaded this session.
    if (allTime.length) {
      await supabase.from('tiktok_reports').delete().eq('source', 'pub_alltime');
      await upsertReports(allTime.slice(0, 500).map((r, i) => toDbRow(r, 'pub_alltime', i) as Record<string, unknown>));
    }
    if (lastMonth.length) {
      await supabase.from('tiktok_reports').delete().eq('source', 'pub_lastmonth');
      await upsertReports(lastMonth.map((r, i) => toDbRow(r, 'pub_lastmonth', i) as Record<string, unknown>));
    }
    if (inhouse.length) {
      await supabase.from('tiktok_reports').delete().eq('source', 'pub_inhouse');
      await upsertReports(inhouse.map((r, i) => toDbRow(r, 'pub_inhouse', i) as Record<string, unknown>));
    }

    // Publish aggregations if available
    if (atAgg) {
      await supabase.from('tiktok_hub_settings').upsert({ key: 'pub_at_agg', value: JSON.stringify(atAgg), updated_at: now });
      setPubAtAgg(atAgg);
    }

    // Publish filters and hidden-video list
    const hiddenArr = JSON.stringify(Array.from(hiddenIds));
    await supabase.from('tiktok_hub_settings').upsert([
      { key: 'pub_filter_at', value: String(savedAt), updated_at: now },
      { key: 'pub_filter_lm', value: String(savedLm), updated_at: now },
      { key: 'pub_filter_cr', value: String(savedCr), updated_at: now },
      { key: 'pub_hidden_ids', value: hiddenArr, updated_at: now },
    ]);

    // Sync local pub state
    setPubAllTime(allTime.slice(0, 500).map((r, i) => ({...r, id:`pub_alltime_${r.videoId||"x"}_${i+1}`, source:'pub_alltime'})));
    setPubLastMonth(lastMonth.map((r, i) => ({...r, id:`pub_lastmonth_${r.videoId||"x"}_${i+1}`, source:'pub_lastmonth'})));
    setPubInhouse(inhouse.map((r, i)   => ({...r, id:`pub_inhouse_${r.videoId||"x"}_${i+1}`,   source:'pub_inhouse'})));
    setPubAt(savedAt);
    setPubLm(savedLm);
    setPubCr(savedCr);
    setPubHiddenIds(new Set(hiddenIds));
    setPublishing(false);
    sessionStorage.removeItem('rl_data_v1');
  };

  // ── file handling ────────────────────────────────────────────────────────────

  const REQUIRED_CSV_COLS = ['URL', 'Creator', 'Video Revenue', 'Items Sold', 'Views Count'];

  const handleFiles = (files: FileList | null) => {
    if (!files) return;
    Array.from(files).forEach(f => {
      const reader = new FileReader();
      reader.onload = async e => { try {
        setUploading(true);
        setUploadStep("Validating file format…");
        // Validate CSV columns before touching any existing data
        const text = e.target!.result as string;
        const firstLine = text.replace(/^﻿/, "").split('\n')[0] || "";
        const missingCols = REQUIRED_CSV_COLS.filter(col => !firstLine.includes(col));
        if (missingCols.length > 0) {
          setUploading(false);
          setUploadStep("");
          alert(`This file doesn't look like a TikTok report. Missing columns: ${missingCols.join(', ')}.\n\nNo data was changed.`);
          return;
        }

        setUploadStep("Loading saved field edits…");
        // Always fetch the latest saved overrides from the DB before parsing so
        // manually-edited fields are never lost when a CSV is re-uploaded.
        const [
          { data: freshOverrideRows },
          { data: freshSettings },
        ] = await Promise.all([
          supabase.from('tiktok_overrides').select('*').order('updated_at', { ascending: true }),
          supabase.from('tiktok_hub_settings').select('key,value'),
        ]);
        const freshOverrides = new Map<string, Override>();
        freshOverrideRows?.forEach((o: Record<string,string>) => {
          freshOverrides.set(o.report_id, {
            audioHook:    o.audio_hook    || undefined,
            visualHook:   o.visual_hook   || undefined,
            textHook:     o.text_hook     || undefined,
            videoLength:  o.video_length  || undefined,
            cta:          o.cta           || undefined,
            sellingPoints:o.selling_points|| undefined,
            keyIdea:      o.key_idea      || undefined,
          });
        });
        // tiktok_hub_settings ov_* keys (guaranteed-working backup)
        freshSettings?.forEach((s: {key:string, value:string}) => {
          if (s.key.startsWith('ov_')) {
            try {
              const videoId = s.key.slice(3);
              const f: Override = JSON.parse(s.value);
              freshOverrides.set(videoId, { ...(freshOverrides.get(videoId) || {}), ...f });
            } catch {}
          }
        });
        // Also fold in any locally-saved overrides not yet persisted to DB
        localOverridesRef.current.forEach((v, k) => freshOverrides.set(k, v));
        setOverridesMap(freshOverrides);

        // Parse raw CSV values — no overrides baked in so tiktok_reports stays clean
        setUploadStep("Parsing CSV rows…");
        const rawRecs = parseCSV(text, upType);

        if (upType === 'alltime') {
          setUploadStep(`Computing stats from ${rawRecs.length.toLocaleString()} rows…`);
          // Compute aggregations from ALL rows before slicing for DB storage
          const allRecs = applyOverrides(rawRecs, freshOverrides);
          const aggCreators = buildCreators(allRecs, lastMonth, inhouse);
          const aggVisual = buildTopHooks(allRecs, 'visualHook');
          const aggText = buildTopHooks(allRecs, 'textHook');
          const aggAudio = buildTopAudioHooks(allRecs);
          const aggCtas = buildTopHooks(allRecs, 'cta');
          const aggSPs = buildTopSellingPoints(allRecs);
          const newAgg = { creators: aggCreators, visualHooks: aggVisual, textHooks: aggText, audioHooks: aggAudio, ctas: aggCtas, sellingPoints: aggSPs };
          setAtAgg(newAgg);
          await supabase.from('tiktok_hub_settings').upsert({ key: 'at_agg', value: JSON.stringify(newAgg), updated_at: new Date().toISOString() });
        }

        // Store raw CSV values in DB — for alltime, limit to top 5000 rows
        setUploadStep("Saving to database…");
        await supabase.from('tiktok_reports').delete().eq('source', upType);
        const dbRawRecs = upType === 'alltime' ? rawRecs.slice(0, 5000) : rawRecs;
        const dbRows = dbRawRecs.map(r => ({
          id:             r.id,
          source:         r.source,
          video_id:       r.videoId,
          video_link:     r.videoLink,
          creator:        r.creator,
          revenue:        r.revenue,
          items_sold:     r.itemsSold,
          views:          r.views,
          likes:          r.likes,
          comments:       r.comments,
          description:    r.description,
          hashtags:       r.hashtags,
          product:        r.product,
          date_posted:    r.datePosted,
          audio_hook:     r.audioHook,
          selling_points: r.sellingPoints,
          key_idea:       r.keyIdea,
          transcript:     r.transcript,
          rank:           r.rank,
        }));
        const saved = await upsertReports(dbRows);
        if (!saved) { alert('Warning: some video rows may not have saved to the database. Please try uploading again.'); }

        // Apply saved overrides to local state so the UI reflects edits immediately
        // For alltime, limit state to 5000 rows for memory efficiency
        const stateRecs = upType === 'alltime' ? rawRecs.slice(0, 5000) : rawRecs;
        const recs = applyOverrides(stateRecs, freshOverrides);
        if (upType==="alltime")   { setAllTime(recs);   setPageAt(20); }
        if (upType==="lastmonth") { setLastMonth(recs); setPageLm(20); }
        if (upType==="inhouse")   setInhouse(recs);
        const now = new Date();
        const dateStr = `${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}-${now.getFullYear()}`;
        await supabase.from('tiktok_hub_settings').upsert({key:'last_imported',value:dateStr,updated_at:new Date().toISOString()});
        setLastImported(dateStr);
        setUploading(false);
        setUploadStep("");
        setShowUp(false);
        sessionStorage.removeItem('rl_data_v1');
      } catch (err) {
        setUploading(false);
        setUploadStep("");
        alert('Upload failed unexpectedly. No data was changed. Please try again.');
        console.error('handleFiles error:', err);
      } };
      reader.readAsText(f);
    });
  };

  // ── card edit panel ──────────────────────────────────────────────────────────

  const openEdit = (r: VideoRow) => {
    setEditingId(r.id);
  };

  const saveEdit = async (r: VideoRow, draft: EditDraft) => {
    const sp = draft.sellingPoints.split("\n").map(s=>s.trim()).filter(Boolean).join(" | ");
    const fields: Override = {
      audioHook:    draft.audioHook,
      visualHook:   draft.visualHook,
      textHook:     draft.textHook,
      videoLength:  draft.videoLength,
      cta:          draft.cta,
      sellingPoints: sp,
      keyIdea:      draft.keyIdea,
    };
    const overrideRow = {
      report_id:     r.videoId,
      audio_hook:    fields.audioHook,
      visual_hook:   fields.visualHook,
      text_hook:     fields.textHook,
      video_length:  fields.videoLength,
      cta:           fields.cta,
      selling_points:fields.sellingPoints,
      key_idea:      fields.keyIdea,
      updated_at:    new Date().toISOString(),
    };
    const { error: upsertErr } = await supabase
      .from('tiktok_overrides').upsert(overrideRow, { onConflict: 'report_id' });
    if (upsertErr) {
      await supabase.from('tiktok_overrides').delete().eq('report_id', r.videoId);
      await supabase.from('tiktok_overrides').insert(overrideRow);
    }
    // Belt-and-suspenders: also persist to tiktok_hub_settings which is confirmed
    // writable. This survives even if tiktok_overrides has RLS or schema issues.
    await supabase.from('tiktok_hub_settings').upsert({
      key: `ov_${r.videoId}`,
      value: JSON.stringify(fields),
      updated_at: new Date().toISOString(),
    });
    // Track in ref so any subsequent load() call re-merges these rather than reverting them
    localOverridesRef.current.set(r.videoId, fields);
    setOverridesMap(prev => new Map(prev).set(r.videoId, fields));
    // Apply to every row with the same videoId across all pages
    const upd = (rs: VideoRow[]) => rs.map(rec => rec.videoId===r.videoId ? {...rec,...fields} : rec);
    const updatedAllTime = allTime.map(rec => rec.videoId===r.videoId ? {...rec,...fields} : rec);
    setAllTime(updatedAllTime); setLastMonth(upd); setInhouse(upd);

    // Recompute hook/CTA/SP sections from updated rows (creators preserved from full-dataset agg)
    if (atAgg) {
      const newAgg = {
        ...atAgg,
        visualHooks:   buildTopHooks(updatedAllTime, 'visualHook'),
        textHooks:     buildTopHooks(updatedAllTime, 'textHook'),
        audioHooks:    buildTopAudioHooks(updatedAllTime),
        ctas:          buildTopHooks(updatedAllTime, 'cta'),
        sellingPoints: buildTopSellingPoints(updatedAllTime),
      };
      setAtAgg(newAgg);
      supabase.from('tiktok_hub_settings').upsert({ key: 'at_agg', value: JSON.stringify(newAgg), updated_at: new Date().toISOString() });
    }

    setEditingId(null);
  };

  const cancelEdit = () => setEditingId(null);
  const toggleTranscript = (id: string) => setTranscriptOpen(prev => { const s=new Set(prev); s.has(id)?s.delete(id):s.add(id); return s; });

  const onAddVisualHookOption = (v: string) => {
    setVisualHookOptions(prev => {
      if (prev.includes(v)) return prev;
      const next = [...prev, v];
      supabase.from('tiktok_hub_settings').upsert({
        key: 'visual_hook_options',
        value: JSON.stringify(next),
        updated_at: new Date().toISOString(),
      });
      return next;
    });
  };

  const importXLSX = (file: File) => {
    setUploading(true);
    setUploadStep("Reading XLSX file…");
    const reader = new FileReader();
    reader.onload = async e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target!.result as ArrayBuffer), { type: 'array' });
        // Collect editable fields keyed by videoId across all video sheets
        const edits = new Map<string, Override>();
        for (const sheetName of ['All-Time Affiliate', 'Last Month Affiliate', 'In-House Affiliate']) {
          const ws = wb.Sheets[sheetName];
          if (!ws) continue;
          const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: '' });
          rows.forEach(row => {
            const s = (v: unknown) => String(v || '').trim();
            const url = s(row['Video URL']);
            const videoId = (url.match(/\/video\/(\d+)/) || [])[1];
            if (!videoId) return;
            const f: Override = {};
            if (s(row['Visual Hook']))    f.visualHook    = s(row['Visual Hook']);
            if (s(row['Text Hook']))      f.textHook      = s(row['Text Hook']);
            if (s(row['Audio Hook']))     f.audioHook     = s(row['Audio Hook']);
            if (s(row['Video Length']))   f.videoLength   = s(row['Video Length']);
            if (s(row['CTA']))            f.cta           = s(row['CTA']);
            if (s(row['Selling Points'])) f.sellingPoints = s(row['Selling Points']);
            if (Object.keys(f).length > 0) {
              edits.set(videoId, { ...(edits.get(videoId) || {}), ...f });
            }
          });
        }
        if (edits.size === 0) {
          setUploading(false);
          alert('No editable fields found. Fill in Visual Hook, Text Hook, Audio Hook, Video Length, CTA, or Selling Points columns and try again.');
          return;
        }
        setUploadStep(`Saving ${edits.size} video override${edits.size !== 1 ? 's' : ''}…`);
        const newMap = new Map(overridesMap);
        for (const [videoId, fields] of Array.from(edits)) {
          const merged: Override = { ...(newMap.get(videoId) || {}), ...fields };
          newMap.set(videoId, merged);
          localOverridesRef.current.set(videoId, merged);
          const overrideRow = { report_id: videoId, audio_hook: merged.audioHook, visual_hook: merged.visualHook, text_hook: merged.textHook, video_length: merged.videoLength, cta: merged.cta, selling_points: merged.sellingPoints, key_idea: merged.keyIdea, updated_at: new Date().toISOString() };
          const { error } = await supabase.from('tiktok_overrides').upsert(overrideRow, { onConflict: 'report_id' });
          if (error) { await supabase.from('tiktok_overrides').delete().eq('report_id', videoId); await supabase.from('tiktok_overrides').insert(overrideRow); }
          await supabase.from('tiktok_hub_settings').upsert({ key: `ov_${videoId}`, value: JSON.stringify(merged), updated_at: new Date().toISOString() });
        }
        setOverridesMap(newMap);
        const applyMap = (rows: VideoRow[]) => rows.map(r => {
          const ov = newMap.get(r.videoId);
          if (!ov) return r;
          return { ...r, audioHook: ov.audioHook ?? r.audioHook, visualHook: ov.visualHook ?? r.visualHook, textHook: ov.textHook ?? r.textHook, videoLength: ov.videoLength ?? r.videoLength, cta: ov.cta ?? r.cta, sellingPoints: ov.sellingPoints ?? r.sellingPoints, keyIdea: ov.keyIdea ?? r.keyIdea };
        });
        const updatedAllTime = applyMap(allTime);
        setAllTime(updatedAllTime); setLastMonth(applyMap(lastMonth)); setInhouse(applyMap(inhouse));
        if (atAgg) {
          const newAgg = { ...atAgg, visualHooks: buildTopHooks(updatedAllTime, 'visualHook'), textHooks: buildTopHooks(updatedAllTime, 'textHook'), audioHooks: buildTopAudioHooks(updatedAllTime), ctas: buildTopHooks(updatedAllTime, 'cta'), sellingPoints: buildTopSellingPoints(updatedAllTime) };
          setAtAgg(newAgg);
          supabase.from('tiktok_hub_settings').upsert({ key: 'at_agg', value: JSON.stringify(newAgg), updated_at: new Date().toISOString() });
        }
        setUploading(false);
        alert(`✓ Updated ${edits.size} video${edits.size !== 1 ? 's' : ''} from XLSX. Changes are live and saved.`);
      } catch {
        setUploading(false);
        alert('Could not read the file. Make sure it was exported from this tool.');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // ── admin filter ─────────────────────────────────────────────────────────────

  const toggleHide = async (videoId: string) => {
    setHiddenIds(prev => {
      const next = new Set(prev);
      if (next.has(videoId)) {
        next.delete(videoId);
        supabase.from('tiktok_hidden_videos').delete().eq('video_id', videoId).then(()=>{});
      } else {
        next.add(videoId);
        supabase.from('tiktok_hidden_videos').insert({ video_id: videoId }).then(()=>{});
      }
      return next;
    });
  };

  const CreatorCard = ({c, idx}: {c: CreatorSummary; idx: number}) => {
    const medals = ["🥇","🥈","🥉"];
    const STATS = [
      {l:"Total Videos",       v:fN(c.totalVideos)},
      {l:"Units Sold",         v:fN(c.unitsSold)},
      {l:"GMV",                v:f$(c.gmv), hi:true},
      {l:"Videos w/ GMV",      v:fN(c.videosWithGmv)},
      {l:"Videos Posted Last Month", v:fN(c.videosLastMonth)},
      {l:"w/ GMV Last Month",  v:fN(c.videosWithGmvLastMonth)},
    ];
    return (
      <div style={{background:"#fff",borderRadius:16,border:"1px solid #e5e7eb",boxShadow:"0 1px 4px rgba(0,0,0,0.06)",marginBottom:24,overflow:"hidden"}}>
        <div style={{display:"flex",alignItems:"center",gap:14,padding:"16px 22px",borderBottom:"1px solid #f0f0f0",background:"#fafafa"}}>
          <div style={{fontSize:idx<3?30:20,minWidth:42,textAlign:"center"}}>{medals[idx]||`#${idx+1}`}</div>
          <div style={{flex:1}}>
            <div style={{fontWeight:800,fontSize:18,color:"#111"}}>@{c.creator}</div>
          </div>
          <div style={{fontWeight:800,fontSize:26,color:"#16a34a"}}>{f$(c.gmv)}</div>
        </div>
        <div className="rl-creator-stats" style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)"}}>
          {STATS.map((s,i)=>(
            <div key={s.l} style={{padding:"14px 16px",borderRight:i<5?"1px solid #f0f0f0":"none",borderBottom:"1px solid #f0f0f0"}}>
              <div style={{fontSize:9,color:"#9ca3af",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:5,lineHeight:1.3}}>{s.l}</div>
              <div style={{fontWeight:700,fontSize:17,color:s.hi?"#16a34a":"#111"}}>{s.v}</div>
            </div>
          ))}
        </div>
        {c.top3.length>0 && (
          <div style={{padding:"16px 22px"}}>
            <div style={{fontSize:11,fontWeight:700,color:"#374151",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:14}}>
              🏆 Top {c.top3.length} Video{c.top3.length!==1?"s":""}
            </div>
            <div className="rl-creator-vid-row" style={{display:"flex",flexWrap:"wrap",gap:16}}>
              {c.top3.map((v,i)=>(
                <div key={v.id||i} className="rl-tile" style={{width:325,flexShrink:0}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                    <span style={{background:"#111",color:"#fff",borderRadius:6,padding:"3px 8px",fontSize:11,fontWeight:700}}>
                      {["1st","2nd","3rd"][i]||`#${i+1}`}
                    </span>
                    <span style={{fontWeight:700,color:"#16a34a",fontSize:13}}>{f$(v.revenue)}</span>
                    <span style={{fontSize:11,color:"#9ca3af"}}>{fN(v.itemsSold)} sold</span>
                    {v.datePosted && <span style={{fontSize:11,color:"#9ca3af"}}>· {v.datePosted}</span>}
                  </div>
                  {v.videoId ? (
                    <div className="rl-tile-embed" style={{width:325,height:578,overflow:"hidden",borderRadius:10,background:"#0a0a0a"}}>
                      <iframe className="rl-tile-iframe" src={`https://www.tiktok.com/embed/v2/${v.videoId}`}
                        style={{display:"block",width:325,height:738,border:"none"}}
                        allowFullScreen allow="encrypted-media" loading="lazy" title={`Top video ${i+1}`}/>
                    </div>
                  ) : (
                    <div className="rl-tile-embed" style={{width:325,height:578,background:"#0a0a0a",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",color:"#555",fontSize:12}}>No embed</div>
                  )}
                  {v.product && (
                    <div style={{marginTop:8,display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                      <span style={{fontSize:9,fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:"0.05em"}}>Product</span>
                      <span style={{fontSize:11,color:"#fff",background:"#374151",borderRadius:20,padding:"2px 9px",fontWeight:500}}>{v.product}</span>
                    </div>
                  )}
                  {pts(v.sellingPoints).length>0 && (
                    <div style={{marginTop:8}}>
                      {pts(v.sellingPoints).slice(0,2).map((p,j)=>(
                        <div key={j} style={{fontSize:11,color:"#4b5563",display:"flex",gap:6,marginBottom:3}}>
                          <span style={{color:"#16a34a",flexShrink:0}}>✓</span>{lbl(p)}
                        </div>
                      ))}
                    </div>
                  )}
                  {(v.visualHook||v.textHook||v.audioHook) && (
                    <div style={{marginTop:8,display:"flex",flexDirection:"column",gap:4}}>
                      {([["🎬","Visual Hook",v.visualHook],["🎣","Text Hook",v.textHook?lbl(pts(v.textHook)[0]||""): ""],[ "🎵","Audio Hook",v.audioHook]] as [string,string,string][]).filter(([,,val])=>val).map(([ic,lb,val])=>(
                        <div key={lb} style={{fontSize:11,color:"#374151",display:"flex",gap:5,alignItems:"flex-start"}}>
                          <span style={{flexShrink:0,fontSize:10}}>{ic}</span>
                          <span style={{fontWeight:600,color:"#6b7280",flexShrink:0}}>{lb}:</span>
                          <span style={{color:"#111",lineHeight:1.4}}>{val}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    );
  };

  // Video tile at original 325px width — no cutoff, stats below
  const VideoTile = ({v, pos}: {v: VideoRow; pos: number}) => (
    <div className="rl-tile" style={{flexShrink:0,width:325}}>
      <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:7,flexWrap:"wrap"}}>
        <span style={{background:"#111",color:"#fff",borderRadius:5,padding:"2px 7px",fontSize:10,fontWeight:700}}>
          {["1st","2nd","3rd"][pos]||`#${pos+1}`}
        </span>
        <span style={{fontWeight:700,color:"#16a34a",fontSize:12}}>{f$(v.revenue)}</span>
        <span style={{fontSize:11,color:"#9ca3af"}}>{fN(v.itemsSold)} sold</span>
        <span style={{fontSize:11,color:"#6b7280",fontWeight:600}}>@{v.creator}</span>
      </div>
      {v.videoId ? (
        <div className="rl-tile-embed" style={{width:325,height:578,overflow:"hidden",borderRadius:10,background:"#0a0a0a"}}>
          <iframe className="rl-tile-iframe" src={`https://www.tiktok.com/embed/v2/${v.videoId}`}
            style={{display:"block",width:325,height:738,border:"none"}}
            allowFullScreen allow="encrypted-media" loading="lazy" title={`@${v.creator}`}/>
        </div>
      ) : (
        <div className="rl-tile-embed" style={{width:325,height:578,background:"#111",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",color:"#555",fontSize:12}}>No embed</div>
      )}
      <div style={{marginTop:8,display:"flex",gap:14}}>
        {[["👁",fK(v.views)],["❤️",fK(v.likes)],["💬",fK(v.comments)]].map(([ic,val])=>(
          <div key={ic as string}>
            <div style={{fontSize:9,color:"#9ca3af",textTransform:"uppercase",letterSpacing:"0.04em",marginBottom:1}}>{ic}</div>
            <div style={{fontWeight:700,fontSize:13,color:"#111"}}>{val}</div>
          </div>
        ))}
      </div>
    </div>
  );

  // One ranked hook entry: stats row + expand button for videos
  const HookEntry = ({h, rank, accent}: {h: HookSummary; rank: number; accent: string}) => {
    const medals = ["🥇","🥈","🥉"];
    const [expanded, setExpanded] = React.useState(false);
    return (
      <div style={{background:"#fff",borderRadius:12,border:"1px solid #e5e7eb",overflow:"hidden",marginBottom:12}}>
        {/* Title row */}
        <div className="rl-hook-entry-hdr" style={{display:"flex",alignItems:"center",gap:10,padding:"11px 16px",background:"#fafafa",borderBottom:"1px solid #f0f0f0"}}>
          <span style={{fontSize:rank<3?22:15,flexShrink:0}}>{medals[rank]||`#${rank+1}`}</span>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontWeight:700,fontSize:13,color:"#111",lineHeight:1.35}}>{h.hookText}</div>
          </div>
          <div style={{fontWeight:800,fontSize:16,color:"#16a34a",flexShrink:0}}>{f$(h.totalGmv)}</div>
        </div>
        {/* Stats row */}
        <div className="rl-stats" style={{display:"flex",gap:0,borderBottom:"1px solid #f0f0f0"}}>
          {[
            ["📹","Total Videos",    String(h.totalVideos)],
            ["👁", "Total Views",     fK(h.totalViews)],
            ["📦","Units Sold",      fN(h.totalUnitsSold)],
            ["💰","Total GMV",       f$(h.totalGmv)],
            ["✅","Videos w/ GMV",  String(h.videosWithGmv)],
          ].map(([ic,label,val],i,arr) => (
            <div key={label} style={{flex:1,padding:"10px 12px",borderRight:i<arr.length-1?"1px solid #f0f0f0":"none",textAlign:"center"}}>
              <div style={{fontSize:9,color:"#9ca3af",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:3}}>{ic} {label}</div>
              <div style={{fontWeight:700,fontSize:13,color:"#111"}}>{val}</div>
            </div>
          ))}
        </div>
        {/* Expand button */}
        {h.topVideos.length>0 && (
          <div style={{padding:"10px 16px",display:"flex",alignItems:"center",gap:8}}>
            <button onClick={()=>setExpanded(e=>!e)}
              style={{display:"flex",alignItems:"center",gap:6,padding:"6px 14px",borderRadius:7,border:"1px solid",
                borderColor:expanded?accent+"66":"#d1d5db",
                background:expanded?accent+"11":"#fff",
                color:expanded?accent:"#374151",
                cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600}}>
              {expanded ? "▲ Hide Videos" : "▼ Show Top Videos"}
            </button>
            {!expanded && <span style={{fontSize:11,color:"#9ca3af"}}>{h.topVideos.length} video{h.topVideos.length!==1?"s":""} available</span>}
          </div>
        )}
        {expanded && h.topVideos.length>0 && (
          <div style={{padding:"0 16px 16px"}}>
            <div style={{fontSize:10,fontWeight:700,color:accent,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:12}}>
              🏆 Top Video{h.topVideos.length>1?"s":""}
            </div>
            <div className="rl-inline-vid-row" style={{display:"flex",gap:16,flexWrap:"wrap"}}>
              {h.topVideos.map((v,i) => <VideoTile key={v.id||i} v={v} pos={i}/>)}
            </div>
          </div>
        )}
      </div>
    );
  };

  // Hook section: accent header + stacked hook entries
  const HookSection = ({hooks, icon, title, accent}: {hooks: HookSummary[]; icon: string; title: string; accent: string}) => {
    if (hooks.length === 0) return null;
    return (
      <div style={{marginBottom:24}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
          <div style={{background:accent,borderRadius:9,width:32,height:32,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0}}>{icon}</div>
          <div style={{fontWeight:800,fontSize:15,color:"#111"}}>{title}</div>
          <div style={{fontSize:11,color:"#9ca3af",marginLeft:2}}>— top {hooks.length} by cumulative GMV</div>
        </div>
        {hooks.map((h,i) => <HookEntry key={h.hookText} h={h} rank={i} accent={accent}/>)}
      </div>
    );
  };

  // Selling point: stats row + expand button for videos
  const SellingPointRow = ({sp, rank}: {sp: SellingPointSummary; rank: number}) => {
    const medals = ["🥇","🥈","🥉","4th","5th"];
    const [expanded, setExpanded] = React.useState(false);
    return (
      <div style={{background:"#fff",borderRadius:12,border:"1px solid #e5e7eb",overflow:"hidden",marginBottom:12}}>
        {/* Title row */}
        <div style={{display:"flex",alignItems:"center",gap:10,padding:"11px 16px",background:"#fafafa",borderBottom:"1px solid #f0f0f0"}}>
          <span style={{fontSize:rank<3?22:14,flexShrink:0,fontWeight:700,color:"#6b7280"}}>{medals[rank]||`#${rank+1}`}</span>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontWeight:700,fontSize:13,color:"#111",lineHeight:1.35}}>{sp.point}</div>
            {sp.product && <span style={{fontSize:10,color:"#fff",background:"#374151",borderRadius:20,padding:"2px 8px",fontWeight:500,marginTop:3,display:"inline-block"}}>{sp.product}</span>}
          </div>
          <div style={{fontWeight:800,fontSize:16,color:"#16a34a",flexShrink:0}}>{f$(sp.totalGmv)}</div>
        </div>
        {/* Stats row */}
        <div className="rl-stats" style={{display:"flex",gap:0,borderBottom:"1px solid #f0f0f0"}}>
          {[
            ["📹","Total Videos",    String(sp.totalVideos)],
            ["👁", "Total Views",     fK(sp.totalViews)],
            ["📦","Units Sold",      fN(sp.totalUnitsSold)],
            ["💰","Total GMV",       f$(sp.totalGmv)],
            ["✅","Videos w/ GMV",  String(sp.videosWithGmv)],
          ].map(([ic,label,val],i,arr) => (
            <div key={label} style={{flex:1,padding:"10px 12px",borderRight:i<arr.length-1?"1px solid #f0f0f0":"none",textAlign:"center"}}>
              <div style={{fontSize:9,color:"#9ca3af",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:3}}>{ic} {label}</div>
              <div style={{fontWeight:700,fontSize:13,color:"#111"}}>{val}</div>
            </div>
          ))}
        </div>
        {/* Expand button */}
        {sp.topVideos.length>0 && (
          <div style={{padding:"10px 16px",display:"flex",alignItems:"center",gap:8}}>
            <button onClick={()=>setExpanded(e=>!e)}
              style={{display:"flex",alignItems:"center",gap:6,padding:"6px 14px",borderRadius:7,border:"1px solid",
                borderColor:expanded?"#16a34a66":"#d1d5db",
                background:expanded?"#f0fdf4":"#fff",
                color:expanded?"#16a34a":"#374151",
                cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600}}>
              {expanded ? "▲ Hide Videos" : "▼ Show Top Videos"}
            </button>
            {!expanded && <span style={{fontSize:11,color:"#9ca3af"}}>{sp.topVideos.length} video{sp.topVideos.length!==1?"s":""} available</span>}
          </div>
        )}
        {expanded && sp.topVideos.length>0 && (
          <div style={{padding:"0 16px 16px"}}>
            <div style={{fontSize:10,fontWeight:700,color:"#16a34a",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:12}}>
              🏆 Top Video{sp.topVideos.length>1?"s":""}
            </div>
            <div className="rl-inline-vid-row" style={{display:"flex",gap:16,flexWrap:"wrap"}}>
              {sp.topVideos.map((v,i) => <VideoTile key={v.id||i} v={v} pos={i}/>)}
            </div>
          </div>
        )}
      </div>
    );
  };

  const slicedAt = visAllTime.slice(0, pageAt);
  const slicedLm = filteredLastMonth.slice(0, pageLm);
  const tabCount = {alltime:visAllTime.length, lastmonth:filteredLastMonth.length, inhouse:visInhouse.length, creators:filteredCreators.length, hooks:topVisualHooks.length+topTextHooks.length+topAudioHooks.length+topCTAs.length+topSellingPoints.length};
  const cardProps = { hiddenIds, editingId, adminMode, transcriptOpen,
    visualHookOptions, toggleHide, cancelEdit, openEdit, saveEdit, toggleTranscript,
    onAddVisualHookOption };
  const gmvAt = visAllTime.reduce((s,r)=>s+r.revenue,0);
  const gmvLm = filteredLastMonth.reduce((s,r)=>s+r.revenue,0);

  const Empty = ({msg}: {msg: string}) => (
    <div style={{textAlign:"center",padding:"60px 20px",color:"#9ca3af"}}>
      <div style={{fontSize:44,marginBottom:12}}>📭</div>
      <div style={{fontSize:15,fontWeight:600,marginBottom:6,color:"#6b7280"}}>No data yet</div>
      <div style={{fontSize:13}}>{msg}</div>
    </div>
  );

  if (dataLoading) {
    return (
      <div style={{minHeight:"100vh",background:"#f4f5f7",display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:12}}>
        <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
        <div style={{width:28,height:28,border:"3px solid #e5e7eb",borderTopColor:"#111",borderRadius:"50%",animation:"spin 0.7s linear infinite"}}/>
        <div style={{color:"#6b7280",fontSize:13}}>Loading reports…</div>
      </div>
    );
  }

  return (
    <div style={{fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif",background:"#f4f5f7",height:"100vh",display:"flex",flexDirection:"column"}}>
      <style>{`
  .rl-tabs{overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none}
  .rl-tabs::-webkit-scrollbar{display:none}
  @media(min-width:900px){
    .rl-two-col{display:grid!important;grid-template-columns:repeat(2,1fr)!important;gap:16px!important;align-items:start!important}
    .rl-two-col>*{margin-bottom:0!important}
    .rl-hooks-grid{display:grid!important;grid-template-columns:1fr 1fr!important;gap:28px!important;align-items:start!important}
    .rl-creator-vid-row,.rl-inline-vid-row{display:grid!important;grid-template-columns:repeat(3,1fr)!important;gap:12px!important}
    .rl-creator-vid-row .rl-tile,.rl-inline-vid-row .rl-tile{width:100%!important;flex-shrink:unset!important}
    .rl-creator-vid-row .rl-tile-embed,.rl-inline-vid-row .rl-tile-embed{width:100%!important;height:auto!important;aspect-ratio:325/578!important}
    .rl-creator-vid-row .rl-tile-iframe,.rl-inline-vid-row .rl-tile-iframe{width:100%!important;height:calc(100% * 738 / 578)!important}
  }
  @media(max-width:639px){
    .rl-header{padding:10px 12px 0!important}
    .rl-hd-date{display:none!important}
    .rl-hd-badges{display:none!important}
    .rl-page{padding:12px!important}
    .rl-subhdr{padding:8px 12px!important;gap:10px!important}
    .rl-card{flex-direction:column!important}
    .rl-card-vid{width:100%!important}
    .rl-embed{width:100%!important;height:calc(100vw * 1.7785)!important}
    .rl-iframe{width:100vw!important;height:calc(100vw * 2.2708)!important}
    .rl-card-body{padding:12px!important}
    .rl-tile{width:72vw!important;max-width:325px}
    .rl-tile-embed{width:72vw!important;max-width:325px;height:calc(72vw * 1.7785)!important;max-height:578px}
    .rl-tile-iframe{width:72vw!important;max-width:325px;height:calc(72vw * 2.2708)!important;max-height:738px}
    .rl-vid-row{flex-wrap:nowrap!important;overflow-x:auto!important;-webkit-overflow-scrolling:touch;padding-bottom:8px;gap:10px!important}
    .rl-stats{overflow-x:auto;-webkit-overflow-scrolling:touch}
    .rl-stats>div{min-width:78px!important;flex-shrink:0!important;flex:none!important}
    .rl-creator-stats{grid-template-columns:repeat(2,1fr)!important}
    .rl-hook-entry-hdr{flex-wrap:wrap!important}
  }
`}</style>

      <div className="rl-header" style={{background:"#0c0c0c",padding:"16px 24px 0",position:"sticky",top:0,zIndex:100,flexShrink:0}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
          <img src="/RuffLinersB.png" alt="Ruff Liners" style={{height:48,width:48,objectFit:"contain",borderRadius:"50%",background:"#fff",flexShrink:0}}/>
          <div>
            <div style={{color:"#fff",fontWeight:800,fontSize:15,letterSpacing:"-0.02em"}}>RUFF LINERS</div>
            <div style={{color:"#a0aec0",fontSize:10,letterSpacing:"0.1em"}}>CREATOR HUB · TIKTOK SHOP</div>
          </div>
          <span style={{flex:1}}/>
          {lastImported && (
            <span className="rl-hd-date" style={{fontSize:11,color:"#e2e8f0",padding:"4px 10px",borderRadius:20,background:"rgba(255,255,255,0.1)",border:"1px solid rgba(255,255,255,0.2)",whiteSpace:"nowrap"}}>
              Date Last Imported: {lastImported}
            </span>
          )}
          {isAdmin && (
            <span className="rl-hd-badges" style={{fontSize:11,fontWeight:700,padding:"4px 12px",borderRadius:20,letterSpacing:"0.05em",
              background:adminMode?"rgba(220,38,38,0.2)":"rgba(255,255,255,0.08)",
              color:adminMode?"#fca5a5":"#9ca3af",border:"1px solid",
              borderColor:adminMode?"rgba(220,38,38,0.4)":"rgba(255,255,255,0.12)"}}>
              {adminMode ? "🔧 ADMIN MODE" : "👁 CREATOR VIEW"}
            </span>
          )}
          {isAdmin && (
            <button onClick={()=>setAdminMode(a=>!a)}
              style={{background:adminMode?"transparent":"#111",color:"#fff",border:"1px solid",
                borderColor:adminMode?"rgba(255,255,255,0.2)":"#333",
                padding:"7px 14px",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600}}>
              {adminMode ? "👁 Exit to Creator View" : "🔧 Enter Admin Mode"}
            </button>
          )}
          {adminMode && (
            <button onClick={()=>setShowUp(u=>!u)}
              style={{background:showUp?"rgba(255,255,255,0.12)":"transparent",color:"#fff",border:"1px solid rgba(255,255,255,0.2)",padding:"7px 14px",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:12}}>
              📁 {showUp?"Close":"Update Reports"}
            </button>
          )}
          {adminMode && (
            <button onClick={()=>buildXLSX(visAllTime, lastMonth, creators)}
              style={{background:"#16a34a",color:"#fff",border:"none",padding:"8px 16px",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600}}>
              📥 Export XLSX
            </button>
          )}
          {adminMode && (
            <button onClick={()=>xlsxImportRef.current?.click()}
              style={{background:"#0891b2",color:"#fff",border:"none",padding:"8px 16px",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600}}>
              📤 Import XLSX
            </button>
          )}
          {adminMode && (
            <button onClick={()=>setShowVHManager(true)}
              style={{background:"#7c3aed",color:"#fff",border:"none",padding:"8px 16px",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600}}>
              🎬 Manage Hooks
            </button>
          )}
          {adminMode && (
            <button onClick={publishDashboard} disabled={publishing}
              style={{background:publishing?"#5b21b6":"#7c3aed",color:"#fff",border:"none",padding:"8px 16px",borderRadius:8,cursor:publishing?"not-allowed":"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600,opacity:publishing?0.8:1}}>
              {publishing?"⏳ Publishing…":"🚀 Update Dashboard"}
            </button>
          )}
          {isLoggedIn ? (
            <button onClick={()=>supabase.auth.signOut()}
              style={{background:"transparent",color:"#666",border:"1px solid #333",padding:"7px 14px",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:12}}>
              Sign Out
            </button>
          ) : (
            <button onClick={()=>setShowLoginModal(true)}
              style={{background:"#3b82f6",color:"#fff",border:"none",padding:"7px 16px",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600}}>
              Sign In
            </button>
          )}
        </div>

        {adminMode && (
          <div style={{background:"rgba(220,38,38,0.12)",border:"1px solid rgba(220,38,38,0.25)",borderRadius:8,padding:"9px 14px",marginBottom:12,fontSize:12,color:"#fca5a5",display:"flex",alignItems:"center",gap:10}}>
            <span>🔧</span>
            <span><strong>Admin Mode</strong> — set GMV filters and save them, or use <strong>👁 Visible / 🚫 Hidden</strong> on each card to control what creators see. Click <strong>Exit to Creator View</strong> to preview exactly what creators will see.</span>
            {hiddenIds.size>0&&<span style={{marginLeft:"auto",background:"rgba(220,38,38,0.2)",borderRadius:20,padding:"3px 12px",whiteSpace:"nowrap"}}>{hiddenIds.size} video{hiddenIds.size!==1?"s":""} hidden</span>}
          </div>
        )}

        {showUp && (
          <div style={{background:"rgba(255,255,255,0.05)",borderRadius:10,padding:"14px",marginBottom:14,border:"1px solid rgba(255,255,255,0.1)"}}>
            <div style={{fontSize:10,color:"rgba(255,255,255,0.4)",marginBottom:10,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em"}}>Upload Euka CSV</div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              <select value={upType} onChange={e=>setUpType(e.target.value)}
                style={{fontSize:12,padding:"8px 10px",borderRadius:7,border:"1px solid rgba(255,255,255,0.2)",background:"rgba(0,0,0,0.5)",color:"#fff",fontFamily:"inherit",minWidth:360}}>
                {UP_TYPES.map(u=><option key={u.value} value={u.value}>{u.label}</option>)}
              </select>
              <div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)}
                onDrop={e=>{e.preventDefault();setDrag(false);handleFiles(e.dataTransfer.files);}}
                onClick={()=>fileRef.current?.click()}
                style={{flex:1,minWidth:200,border:`1.5px dashed ${drag?"#fff":"rgba(255,255,255,0.25)"}`,borderRadius:7,padding:"8px 14px",cursor:"pointer",display:"flex",alignItems:"center",gap:8}}>
                <span>📁</span>
                <span style={{fontSize:12,color:"rgba(255,255,255,0.5)"}}>Drop CSV or click to browse — your notes carry over automatically</span>
                <input ref={fileRef} type="file" accept=".csv" multiple style={{display:"none"}} onChange={e=>handleFiles(e.target.files)}/>
              </div>
            </div>
          </div>
        )}

        {/* Hidden XLSX import input — kept outside showUp so it's always mounted */}
        <input ref={xlsxImportRef} type="file" accept=".xlsx" style={{display:"none"}} onChange={e=>{ const f=e.target.files?.[0]; if(f) importXLSX(f); e.target.value=''; }}/>

        <div className="rl-tabs" style={{display:"flex",overflowX:"auto"}}>
          {TABS.map(t=>{
            const active=tab===t.id;
            return (
              <button key={t.id} onClick={()=>setTab(t.id)}
                style={{padding:"12px 18px",border:"none",borderBottom:active?"3px solid #fff":"3px solid transparent",background:"none",cursor:"pointer",fontSize:13,fontWeight:active?700:500,color:active?"#fff":"#b0bec5",whiteSpace:"nowrap",fontFamily:"inherit",display:"flex",alignItems:"center",gap:7,transition:"color .12s"}}>
                {t.icon} {t.label}
                <span style={{background:active?"rgba(255,255,255,0.14)":"rgba(255,255,255,0.08)",color:active?"#fff":"#90a4b4",borderRadius:20,padding:"2px 8px",fontSize:10,fontWeight:700}}>
                  {tabCount[t.id as keyof typeof tabCount]}
                </span>
              </button>
            );
          })}
        </div>
      </div>

      {tab==="alltime" && (
        <>
          <div className="rl-subhdr" style={{background:"#fff",borderBottom:"1px solid #e5e7eb",padding:"10px 24px",display:"flex",gap:20,alignItems:"center",flexWrap:"wrap"}}>
            {visAllTime.length>0 && <>
              <span style={{fontSize:12,color:"#9ca3af"}}><b style={{color:"#111"}}>{visAllTime.length}</b> videos</span>
              <span style={{fontSize:12,color:"#9ca3af"}}>GMV: <b style={{color:"#16a34a"}}>{f$(gmvAt)}</b></span>
              <span style={{fontSize:12,color:"#9ca3af"}}>Views: <b style={{color:"#2563eb"}}>{fK(visAllTime.reduce((s,r)=>s+r.views,0))}</b></span>
            </>}
            {adminMode && (
              <span style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:8,fontSize:12,color:"#6b7280",flexShrink:0}}>
                Min GMV:
                <span style={{display:"inline-flex",alignItems:"center",gap:2}}>
                  <span style={{color:"#9ca3af"}}>$</span>
                  <input type="number" min="0" step="100" value={draftAt} placeholder="0"
                    onChange={e=>setDraftAt(e.target.value as unknown as number)}
                    onKeyDown={e=>e.key==="Enter"&&saveFilterAt()}
                    style={{width:90,padding:"4px 8px",border:`1px solid ${Number(draftAt)!==savedAt?"#f59e0b":"#d1d5db"}`,borderRadius:6,fontFamily:"inherit",fontSize:12,color:"#111",outline:"none"}}/>
                </span>
                <button onClick={saveFilterAt}
                  style={{padding:"4px 12px",background:Number(draftAt||0)!==savedAt?"#111":"#f3f4f6",color:Number(draftAt||0)!==savedAt?"#fff":"#9ca3af",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600,transition:"all .15s"}}>
                  Save
                </button>
                {savedAt>0&&<button onClick={clearFilterAt} style={{fontSize:11,color:"#9ca3af",background:"none",border:"none",cursor:"pointer",padding:0,fontFamily:"inherit"}}>Clear</button>}
              </span>
            )}
            {adminMode && hiddenIds.size>0 && <span style={{fontSize:11,color:"#dc2626",flexShrink:0}}>🚫 {hiddenIds.size} hidden</span>}
          </div>
          {(adminMode ? savedAt : pubAt)>0 && (
            <div style={{background:"#111",padding:"9px 24px",display:"flex",alignItems:"center",gap:10}}>
              <span style={{fontSize:14}}>🏆</span>
              <span style={{color:"#fff",fontSize:13,fontWeight:600}}>All Time Videos with GMV &gt; {f$(adminMode ? savedAt : pubAt)}</span>
            </div>
          )}
        </>
      )}

      {tab==="lastmonth" && (
        <>
          <div className="rl-subhdr" style={{background:"#fff",borderBottom:"1px solid #e5e7eb",padding:"10px 24px",display:"flex",gap:20,alignItems:"center",flexWrap:"wrap"}}>
            {filteredLastMonth.length>0 && <>
              <span style={{fontSize:12,color:"#9ca3af"}}><b style={{color:"#111"}}>{filteredLastMonth.length}</b> videos</span>
              <span style={{fontSize:12,color:"#9ca3af"}}>GMV: <b style={{color:"#16a34a"}}>{f$(gmvLm)}</b></span>
              <span style={{fontSize:12,color:"#9ca3af"}}>Views: <b style={{color:"#2563eb"}}>{fK(filteredLastMonth.reduce((s,r)=>s+r.views,0))}</b></span>
            </>}
            {adminMode && (
              <span style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:8,fontSize:12,color:"#6b7280",flexShrink:0}}>
                Min GMV:
                <span style={{display:"inline-flex",alignItems:"center",gap:2}}>
                  <span style={{color:"#9ca3af"}}>$</span>
                  <input type="number" min="0" step="100" value={draftLm} placeholder="0"
                    onChange={e=>setDraftLm(e.target.value as unknown as number)}
                    onKeyDown={e=>e.key==="Enter"&&saveFilterLm()}
                    style={{width:90,padding:"4px 8px",border:`1px solid ${Number(draftLm)!==savedLm?"#f59e0b":"#d1d5db"}`,borderRadius:6,fontFamily:"inherit",fontSize:12,color:"#111",outline:"none"}}/>
                </span>
                <button onClick={saveFilterLm}
                  style={{padding:"4px 12px",background:Number(draftLm||0)!==savedLm?"#111":"#f3f4f6",color:Number(draftLm||0)!==savedLm?"#fff":"#9ca3af",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600,transition:"all .15s"}}>
                  Save
                </button>
                {savedLm>0&&<button onClick={clearFilterLm} style={{fontSize:11,color:"#9ca3af",background:"none",border:"none",cursor:"pointer",padding:0,fontFamily:"inherit"}}>Clear</button>}
              </span>
            )}
          </div>
          {(adminMode ? savedLm : pubLm)>0 && (
            <div style={{background:"#111",padding:"9px 24px",display:"flex",alignItems:"center",gap:10}}>
              <span style={{fontSize:14}}>📅</span>
              <span style={{color:"#fff",fontSize:13,fontWeight:600}}>Last Month Videos with GMV &gt; {f$(adminMode ? savedLm : pubLm)}</span>
            </div>
          )}
        </>
      )}

      {tab==="creators" && (
        <>
          <div className="rl-subhdr" style={{background:"#fff",borderBottom:"1px solid #e5e7eb",padding:"10px 24px",display:"flex",gap:20,alignItems:"center",flexWrap:"wrap"}}>
            {filteredCreators.length>0 && <>
              <span style={{fontSize:12,color:"#9ca3af"}}><b style={{color:"#111"}}>{filteredCreators.length}</b> creators</span>
              <span style={{fontSize:12,color:"#9ca3af"}}>Combined GMV: <b style={{color:"#16a34a"}}>{f$(filteredCreators.reduce((s,c)=>s+c.gmv,0))}</b></span>
              <span style={{fontSize:12,color:"#9ca3af"}}>Total Videos: <b style={{color:"#2563eb"}}>{fN(filteredCreators.reduce((s,c)=>s+c.totalVideos,0))}</b></span>
            </>}
            {adminMode && (
              <span style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:8,fontSize:12,color:"#6b7280",flexShrink:0}}>
                Min Creator GMV:
                <span style={{display:"inline-flex",alignItems:"center",gap:2}}>
                  <span style={{color:"#9ca3af"}}>$</span>
                  <input type="number" min="0" step="1000" value={draftCr} placeholder="0"
                    onChange={e=>setDraftCr(e.target.value as unknown as number)}
                    onKeyDown={e=>e.key==="Enter"&&saveFilterCr()}
                    style={{width:100,padding:"4px 8px",border:`1px solid ${Number(draftCr)!==savedCr?"#f59e0b":"#d1d5db"}`,borderRadius:6,fontFamily:"inherit",fontSize:12,color:"#111",outline:"none"}}/>
                </span>
                <button onClick={saveFilterCr}
                  style={{padding:"4px 12px",background:Number(draftCr||0)!==savedCr?"#111":"#f3f4f6",color:Number(draftCr||0)!==savedCr?"#fff":"#9ca3af",border:"none",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:600,transition:"all .15s"}}>
                  Save
                </button>
                {savedCr>0&&<button onClick={clearFilterCr} style={{fontSize:11,color:"#9ca3af",background:"none",border:"none",cursor:"pointer",padding:0,fontFamily:"inherit"}}>Clear</button>}
              </span>
            )}
          </div>
          {(adminMode ? savedCr : pubCr)>0 && (
            <div style={{background:"#111",padding:"9px 24px",display:"flex",alignItems:"center",gap:10}}>
              <span style={{fontSize:14}}>⭐</span>
              <span style={{color:"#fff",fontSize:13,fontWeight:600}}>Top Creators with Total GMV &gt; {f$(adminMode ? savedCr : pubCr)}</span>
            </div>
          )}
        </>
      )}

      <div style={{flex:1,overflowY:"auto"}}>

      <div className="rl-page" style={{padding:"20px 24px"}}>

        {tab==="alltime" && (
          visAllTime.length===0
            ? <Empty msg='Click "Update Reports" above and upload the All-Time export from app.euka.ai/videos'/>
            : <div className="rl-two-col">
                {slicedAt.map(r=><VideoCard key={r.id} r={r} showFilter={adminMode} {...cardProps}/>)}
                {pageAt<visAllTime.length && (
                  <div style={{gridColumn:"1/-1",textAlign:"center",padding:"16px 0"}}>
                    <button onClick={()=>setPageAt(n=>n+20)}
                      style={{padding:"10px 28px",background:"#fff",border:"1px solid #d1d5db",borderRadius:10,cursor:"pointer",fontFamily:"inherit",fontSize:13,color:"#374151",fontWeight:500}}>
                      Show more ({visAllTime.length-pageAt} remaining)
                    </button>
                  </div>
                )}
              </div>
        )}

        {tab==="lastmonth" && (
          filteredLastMonth.length===0
            ? <Empty msg='Click "Update Reports" above and upload the Last Month export from app.euka.ai/videos'/>
            : <div className="rl-two-col">
                {slicedLm.map(r=><VideoCard key={r.id} r={r} showFilter={false} {...cardProps}/>)}
                {pageLm<filteredLastMonth.length && (
                  <div style={{gridColumn:"1/-1",textAlign:"center",padding:"16px 0"}}>
                    <button onClick={()=>setPageLm(n=>n+20)}
                      style={{padding:"10px 28px",background:"#fff",border:"1px solid #d1d5db",borderRadius:10,cursor:"pointer",fontFamily:"inherit",fontSize:13,color:"#374151",fontWeight:500}}>
                      Show more ({filteredLastMonth.length-pageLm} remaining)
                    </button>
                  </div>
                )}
              </div>
        )}

        {tab==="inhouse" && (
          visInhouse.length===0
            ? <Empty msg="In-House Content report source TBD — upload when available"/>
            : <div className="rl-two-col">{visInhouse.map(r=><VideoCard key={r.id} r={r} showFilter={false} {...cardProps}/>)}</div>
        )}

        {tab==="creators" && (
          filteredCreators.length===0
            ? <Empty msg='Upload the All-Time report first — Top Creators are computed from that data'/>
            : <div className="rl-two-col">{filteredCreators.map((c,i)=><CreatorCard key={c.creator} c={c} idx={i}/>)}</div>
        )}

        {tab==="hooks" && (
          topVisualHooks.length===0 && topTextHooks.length===0 && topAudioHooks.length===0 && topCTAs.length===0 && topSellingPoints.length===0
            ? <Empty msg='Upload the All-Time report first — Hooks, CTAs, and Selling Points are computed from that data'/>
            : <div className="rl-hooks-grid">
                {/* ── Left col: Hooks ── */}
                <div>
                  <div style={{marginBottom:8}}>
                    <div style={{fontWeight:900,fontSize:22,color:"#111",marginBottom:2}}>🪝 Hooks</div>
                    <div style={{fontSize:12,color:"#9ca3af",marginBottom:20}}>Top 3 per type by cumulative GMV across all all-time videos</div>
                  </div>
                  <HookSection hooks={topVisualHooks} icon="🎬" title="Visual Hooks" accent="#7c3aed"/>
                  <HookSection hooks={topTextHooks}   icon="🎣" title="Text Hooks"   accent="#0891b2"/>
                  <HookSection hooks={topAudioHooks}  icon="🎵" title="Audio Hooks"  accent="#d97706"/>
                </div>

                {/* ── Right col: CTAs + Selling Points ── */}
                <div>
                  <div style={{marginBottom:8}}>
                    <div style={{fontWeight:900,fontSize:22,color:"#111",marginBottom:2}}>📣 Call to Action</div>
                    <div style={{fontSize:12,color:"#9ca3af",marginBottom:20}}>Top CTAs by cumulative GMV across all all-time videos</div>
                  </div>
                  {topCTAs.length>0
                    ? <HookSection hooks={topCTAs} icon="📣" title="Top CTAs" accent="#dc2626"/>
                    : <div style={{background:"#fff",borderRadius:12,border:"1px dashed #d1d5db",padding:"28px 20px",textAlign:"center",marginBottom:28,color:"#9ca3af",fontSize:13}}>
                        No CTA data yet — edit videos and fill in the <strong style={{color:"#374151"}}>Call to Action</strong> field to see top CTAs here.
                      </div>
                  }
                  {topSellingPoints.length>0 && (
                    <>
                      <div style={{borderTop:"2px solid #e5e7eb",marginTop:8,paddingTop:28,marginBottom:8}}>
                        <div style={{fontWeight:900,fontSize:22,color:"#111",marginBottom:2}}>✅ Best Selling Points</div>
                        <div style={{fontSize:12,color:"#9ca3af",marginBottom:20}}>Top 5 individual selling points by cumulative GMV, grouped per product</div>
                      </div>
                      <div style={{display:"flex",flexDirection:"column",gap:12}}>
                        {topSellingPoints.map((sp,i)=><SellingPointRow key={`${sp.product}::${sp.point}`} sp={sp} rank={i}/>)}
                      </div>
                    </>
                  )}
                </div>
              </div>
        )}
      </div>

      </div>{/* end scrollable content */}


{uploading && (
  <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center"}}>
    <div style={{background:"#fff",borderRadius:16,padding:"32px 40px",minWidth:300,textAlign:"center",boxShadow:"0 20px 60px rgba(0,0,0,0.25)"}}>
      <style>{`@keyframes rl-spin{to{transform:rotate(360deg)}}`}</style>
      <div style={{width:44,height:44,border:"4px solid #e5e7eb",borderTopColor:"#16a34a",borderRadius:"50%",animation:"rl-spin 0.8s linear infinite",margin:"0 auto 20px"}}/>
      <div style={{fontWeight:700,fontSize:16,color:"#111",marginBottom:8}}>Uploading Report…</div>
      <div style={{fontSize:13,color:"#6b7280"}}>{uploadStep}</div>
    </div>
  </div>
)}

{showLoginModal && (
        <div
          onClick={(e)=>{if(e.target===e.currentTarget)setShowLoginModal(false);}}
          style={{position:"fixed",inset:0,zIndex:2000,background:"rgba(0,0,0,0.75)",display:"flex",alignItems:"center",justifyContent:"center"}}
        >
          <div style={{background:"#1a2035",border:"1px solid #2d3748",borderRadius:12,padding:40,width:360}}>
            <div style={{textAlign:"center",marginBottom:24}}>
              <div style={{fontSize:28}}>🐾</div>
              <h2 style={{color:"#f9fafb",fontSize:18,fontWeight:700,margin:"8px 0 4px"}}>Sign In</h2>
              <p style={{color:"#6b7280",fontSize:13,margin:0}}>Sign in to access admin features.</p>
            </div>
            <form onSubmit={async(e)=>{
              e.preventDefault();
              setLoginLoading(true); setLoginError('');
              const {error} = await supabase.auth.signInWithPassword({email:loginEmail,password:loginPassword});
              if(error){setLoginError(error.message);setLoginLoading(false);}
              else{setShowLoginModal(false);setLoginEmail('');setLoginPassword('');setLoginLoading(false);}
            }} style={{display:"flex",flexDirection:"column",gap:14}}>
              <div>
                <label style={{fontSize:12,fontWeight:600,color:"#9ca3af",display:"block",marginBottom:6}}>EMAIL</label>
                <input type="email" value={loginEmail} onChange={e=>setLoginEmail(e.target.value)} placeholder="you@ruffliners.com" required
                  style={{width:"100%",background:"#0d1117",border:"1px solid #2d3748",borderRadius:6,color:"#e5e7eb",padding:"10px 12px",fontSize:13,boxSizing:"border-box",outline:"none"}}/>
              </div>
              <div>
                <label style={{fontSize:12,fontWeight:600,color:"#9ca3af",display:"block",marginBottom:6}}>PASSWORD</label>
                <input type="password" value={loginPassword} onChange={e=>setLoginPassword(e.target.value)} placeholder="••••••••" required
                  style={{width:"100%",background:"#0d1117",border:"1px solid #2d3748",borderRadius:6,color:"#e5e7eb",padding:"10px 12px",fontSize:13,boxSizing:"border-box",outline:"none"}}/>
              </div>
              {loginError && (
                <div style={{background:"#ef444422",border:"1px solid #ef4444",borderRadius:6,padding:"8px 12px",fontSize:12,color:"#ef4444"}}>{loginError}</div>
              )}
              <button type="submit" disabled={loginLoading}
                style={{background:loginLoading?"#1e3a5f":"#3b82f6",color:"#fff",border:"none",borderRadius:7,padding:"11px 0",fontSize:14,fontWeight:700,cursor:loginLoading?"not-allowed":"pointer",marginTop:4}}>
                {loginLoading?"Please wait…":"Sign In"}
              </button>
            </form>
            <button onClick={()=>setShowLoginModal(false)}
              style={{display:"block",width:"100%",marginTop:12,background:"none",border:"none",color:"#6b7280",fontSize:12,cursor:"pointer",fontFamily:"inherit",padding:"6px 0"}}>
              Cancel
            </button>
          </div>
        </div>
      )}

{showVHManager && (
  <VHManagerModal
    options={visualHookOptions}
    onClose={()=>setShowVHManager(false)}
    onSave={async (next)=>{
      setVisualHookOptions(next);
      supabase.from('tiktok_hub_settings').upsert({
        key: 'visual_hook_options',
        value: JSON.stringify(next),
        updated_at: new Date().toISOString(),
      });
      // Clear visualHook on any video whose current value isn't in the new list
      const allowed = new Set(next);
      const toClear = allTime.filter(r => r.visualHook && !allowed.has(r.visualHook));
      if (toClear.length > 0) {
        const clearOps = toClear.map(row => ({
          report_id: row.videoId,
          visual_hook: null,
          updated_at: new Date().toISOString(),
        }));
        await supabase.from('tiktok_overrides').upsert(clearOps, { onConflict: 'report_id' });
        // Update in-memory state
        const clearIds = new Set(toClear.map(r => r.videoId));
        const clearRow = (rs: VideoRow[]) => rs.map(r => clearIds.has(r.videoId) ? {...r, visualHook:""} : r);
        setAllTime(clearRow); setLastMonth(clearRow); setInhouse(clearRow);
      }
    }}
  />
)}

    </div>
  );
}
