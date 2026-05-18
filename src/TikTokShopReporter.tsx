import { useState, useRef, useMemo, useEffect } from "react";
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

interface Override {
  audioHook?: string;
  visualHook?: string;
  textHook?: string;
  videoLength?: string;
  sellingPoints?: string;
}

interface EditDraft {
  audioHook: string;
  visualHook: string;
  textHook: string;
  videoLength: string;
  sellingPoints: string;
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

const parseCSV = (text: string, source: string, overridesMap: Map<string, Override>): VideoRow[] => {
  const result = Papa.parse<Record<string,string>>(text.replace(/^﻿/,"").trim(), {header:true, skipEmptyLines:true});
  const rows = result.data.filter(r => r.URL && r.Creator);
  const sorted = rows.map(r => {
    const url = (r.URL||"").trim();
    const vid = (url.match(/\/video\/(\d+)/)||[])[1]||"";
    const rev = parseFloat((r["Video Revenue"]||"").replace(/[$,]/g,""))||0;
    const desc = (r.Description||"").replace(/#\w+/g,"").trim();
    const tags = ((r.Description||"").match(/#\w+/g)||[]).join(" ");
    const override = overridesMap.get(`${source}_${vid}`) || {};
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
      audioHook:    override.audioHook    ?? C(r.Hooks||""),
      textHook:     override.textHook     ?? "",
      sellingPoints:override.sellingPoints ?? C(r["Selling Points"]||""),
      visualHook:   override.visualHook   ?? "",
      videoLength:  override.videoLength  ?? "",
      keyIdea:   C(r["Key Idea"] || ""),
      transcript: C(r.Transcript || ""),
    } as Omit<VideoRow, 'id' | 'rank'>;
  }).sort((a,b)=>b.revenue-a.revenue).map((r,i)=>({...r, id:`${source}_${r.videoId||i}`, rank:i+1}));
  return sorted as VideoRow[];
};

// ─── EXPORT ───────────────────────────────────────────────────────────────────

const buildXLSX = (at: VideoRow[], lm: VideoRow[], creators: CreatorSummary[]) => {
  const wb = XLSX.utils.book_new();
  const vH = ["#","Creator","Video URL","Revenue ($)","Items Sold","Views","Likes","Comments","Product","Description","Hashtags","Visual Hook","Text Hook","Audio Hook","Video Length","Selling Points"];
  const vR = (r: VideoRow, i: number): (string|number)[] => [i+1,r.creator,r.videoLink,r.revenue,r.itemsSold,r.views,r.likes,r.comments,r.product,r.description,r.hashtags,r.visualHook,r.textHook,r.audioHook,r.videoLength,r.sellingPoints];
  const mkSheet = (rows: VideoRow[], name: string) => {
    const ws = XLSX.utils.aoa_to_sheet([vH,...rows.map((r,i)=>vR(r,i))]);
    ws["!cols"] = [4,18,30,12,10,12,10,10,28,40,22,22,22,16,12,40].map(w=>({wch:w}));
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
  {id:"alltime",   label:"All-Time High GMV",   icon:"🏆"},
  {id:"lastmonth", label:"Last Month High GMV",  icon:"📅"},
  {id:"inhouse",   label:"In-House High GMV",    icon:"🎬"},
  {id:"creators",  label:"Top Creators",         icon:"⭐"},
];

const UP_TYPES = [
  {value:"alltime",   label:"All-Time Report  →  feeds All-Time tab + Top Creators"},
  {value:"lastmonth", label:"Last Month Report  →  feeds Last Month tab + monthly creator stats"},
  {value:"inhouse",   label:"In-House Content Report"},
];

// ─── COMPONENT ────────────────────────────────────────────────────────────────

export default function TikTokShopReporter() {

  const [tab,       setTab]       = useState("alltime");
  const [allTime,   setAllTime]   = useState<VideoRow[]>([]);
  const [lastMonth, setLastMonth] = useState<VideoRow[]>([]);
  const [inhouse,   setInhouse]   = useState<VideoRow[]>([]);
  const [adminMode, setAdminMode] = useState(false);
  const [isAdmin,   setIsAdmin]   = useState(false);
  const [hiddenIds, setHiddenIds] = useState<Set<string>>(new Set());
  const [editingId,  setEditingId]  = useState<string | null>(null);
  const [editDraft,  setEditDraft]  = useState<EditDraft>({audioHook:"",visualHook:"",textHook:"",videoLength:"",sellingPoints:""});
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
  const [dataLoading, setDataLoading] = useState(true);
  const fileRef = useRef<HTMLInputElement>(null);
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

  // ── initial load from Supabase ──────────────────────────────────────────────

  useEffect(() => {
    const { data: { subscription: authSub } } = supabase.auth.onAuthStateChange((_event, session) => {
      const loggedIn = !!session;
      setIsLoggedIn(loggedIn);
      const adminFlag = session?.user?.user_metadata?.is_admin === true;
      setIsAdmin(adminFlag);
      if (!loggedIn) { setAdminMode(false); setShowLoginModal(false); }
    });

    const load = async () => {
      setDataLoading(true);

      // Detect admin from session metadata
      const { data: { session } } = await supabase.auth.getSession();
      const loggedIn = !!session;
      setIsLoggedIn(loggedIn);
      const adminFlag = session?.user?.user_metadata?.is_admin === true;
      setIsAdmin(adminFlag);

      // Load overrides
      const { data: overrideRows } = await supabase.from('tiktok_overrides').select('*');
      const newMap = new Map<string, Override>();
      overrideRows?.forEach((o: Record<string,string>) => {
        newMap.set(o.report_id, {
          audioHook:    o.audio_hook    || undefined,
          visualHook:   o.visual_hook   || undefined,
          textHook:     o.text_hook     || undefined,
          videoLength:  o.video_length  || undefined,
          sellingPoints:o.selling_points|| undefined,
        });
      });
      setOverridesMap(newMap);

      // Load hidden videos
      const { data: hiddenRows } = await supabase.from('tiktok_hidden_videos').select('video_id');
      setHiddenIds(new Set(hiddenRows?.map((h: {video_id:string}) => h.video_id) || []));

      // Load settings
      const { data: settings } = await supabase.from('tiktok_hub_settings').select('*');
      settings?.forEach((s: {key:string, value:string}) => {
        const n = Number(s.value) || 0;
        if (s.key === 'filter_at') { setSavedAt(n); setDraftAt(n || ""); }
        if (s.key === 'filter_lm') { setSavedLm(n); setDraftLm(n || ""); }
        if (s.key === 'filter_cr') { setSavedCr(n); setDraftCr(n || ""); }
        if (s.key === 'pub_filter_at') setPubAt(Number(s.value) || 0);
        if (s.key === 'pub_filter_lm') setPubLm(Number(s.value) || 0);
        if (s.key === 'pub_filter_cr') setPubCr(Number(s.value) || 0);
        if (s.key === 'pub_hidden_ids') { try { setPubHiddenIds(new Set(JSON.parse(s.value))); } catch {} }
        if (s.key === 'last_imported') setLastImported(s.value);
      });

      // Load reports
      const { data: reports } = await supabase.from('tiktok_reports').select('*');
      if (reports) {
        const toRow = (r: Record<string,unknown>, om: Map<string,Override>): VideoRow => {
          // pub_alltime_xyz → alltime_xyz for override lookup
          const ovKey = (r.id as string).replace(/^pub_/, '');
          const ov = om.get(ovKey) || {};
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
            sellingPoints:ov.sellingPoints ?? r.selling_points as string ?? "",
            keyIdea:      r.key_idea as string ?? "",
            transcript:   r.transcript as string ?? "",
            rank:         r.rank as number,
          };
        };
        const srt = (a: Record<string,unknown>, b: Record<string,unknown>) => (b.revenue as number) - (a.revenue as number);
        setAllTime(reports.filter(r => r.source==='alltime').sort(srt).map(r=>toRow(r,newMap)));
        setLastMonth(reports.filter(r => r.source==='lastmonth').sort(srt).map(r=>toRow(r,newMap)));
        setInhouse(reports.filter(r => r.source==='inhouse').sort(srt).map(r=>toRow(r,newMap)));
        setPubAllTime(reports.filter(r => r.source==='pub_alltime').sort(srt).map(r=>toRow(r,newMap)));
        setPubLastMonth(reports.filter(r => r.source==='pub_lastmonth').sort(srt).map(r=>toRow(r,newMap)));
        setPubInhouse(reports.filter(r => r.source==='pub_inhouse').sort(srt).map(r=>toRow(r,newMap)));
      }

      setDataLoading(false);
    };
    load();
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

  const buildCreators = (atV: VideoRow[], lmV: VideoRow[]): CreatorSummary[] => {
    if (!atV.length) return [];
    const map: Record<string, {creator:string; allV:VideoRow[]; lmV:VideoRow[]}> = {};
    atV.forEach(v => {
      if (!map[v.creator]) map[v.creator] = {creator:v.creator, allV:[], lmV:[]};
      map[v.creator].allV.push(v);
    });
    lmV.forEach(v => {
      if (!map[v.creator]) map[v.creator] = {creator:v.creator, allV:[], lmV:[]};
      map[v.creator].lmV.push(v);
    });
    return Object.values(map)
      .map(c => {
        const top3 = [...c.allV].sort((a,b)=>b.revenue-a.revenue).slice(0,3);
        return {
          creator:               c.creator,
          totalVideos:           c.allV.length,
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
  const creators    = useMemo(() => buildCreators(allTime,    lastMonth),    [allTime,    lastMonth]);     // eslint-disable-line react-hooks/exhaustive-deps
  const pubCreators = useMemo(() => buildCreators(pubAllTime, pubLastMonth), [pubAllTime, pubLastMonth]); // eslint-disable-line react-hooks/exhaustive-deps

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

  const filteredCreators = useMemo(
    () => {
      const src = adminMode ? creators : pubCreators;
      const threshold = adminMode ? savedCr : pubCr;
      return threshold > 0 ? src.filter(c => c.gmv >= threshold) : src;
    },
    [creators, pubCreators, adminMode, savedCr, pubCr]
  );

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

    const toDbRow = (r: VideoRow, pubSource: string) => ({
      id: `pub_${r.id}`,
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

    // Replace all published video snapshots
    await supabase.from('tiktok_reports').delete().in('source', ['pub_alltime','pub_lastmonth','pub_inhouse']);
    if (allTime.length)   await supabase.from('tiktok_reports').insert(allTime.map(r   => toDbRow(r,'pub_alltime')));
    if (lastMonth.length) await supabase.from('tiktok_reports').insert(lastMonth.map(r => toDbRow(r,'pub_lastmonth')));
    if (inhouse.length)   await supabase.from('tiktok_reports').insert(inhouse.map(r   => toDbRow(r,'pub_inhouse')));

    // Publish filters and hidden-video list
    const hiddenArr = JSON.stringify(Array.from(hiddenIds));
    await supabase.from('tiktok_hub_settings').upsert([
      { key: 'pub_filter_at', value: String(savedAt), updated_at: now },
      { key: 'pub_filter_lm', value: String(savedLm), updated_at: now },
      { key: 'pub_filter_cr', value: String(savedCr), updated_at: now },
      { key: 'pub_hidden_ids', value: hiddenArr, updated_at: now },
    ]);

    // Sync local pub state
    setPubAllTime(allTime.map(r   => ({...r, id:`pub_${r.id}`, source:'pub_alltime'})));
    setPubLastMonth(lastMonth.map(r => ({...r, id:`pub_${r.id}`, source:'pub_lastmonth'})));
    setPubInhouse(inhouse.map(r   => ({...r, id:`pub_${r.id}`, source:'pub_inhouse'})));
    setPubAt(savedAt);
    setPubLm(savedLm);
    setPubCr(savedCr);
    setPubHiddenIds(new Set(hiddenIds));
    setPublishing(false);
  };

  // ── file handling ────────────────────────────────────────────────────────────

  const handleFiles = (files: FileList | null) => {
    if (!files) return;
    Array.from(files).forEach(f => {
      const reader = new FileReader();
      reader.onload = async e => {
        const recs = parseCSV(e.target!.result as string, upType, overridesMap);

        // Replace existing rows for this source in Supabase
        await supabase.from('tiktok_reports').delete().eq('source', upType);
        await supabase.from('tiktok_reports').insert(
          recs.map(r => ({
            id:           r.id,
            source:       r.source,
            video_id:     r.videoId,
            video_link:   r.videoLink,
            creator:      r.creator,
            revenue:      r.revenue,
            items_sold:   r.itemsSold,
            views:        r.views,
            likes:        r.likes,
            comments:     r.comments,
            description:  r.description,
            hashtags:     r.hashtags,
            product:      r.product,
            date_posted:  r.datePosted,
            audio_hook:   r.audioHook,
            selling_points: r.sellingPoints,
            key_idea:     r.keyIdea,
            transcript:   r.transcript,
            rank:         r.rank,
          }))
        );

        if (upType==="alltime")   { setAllTime(recs);   setPageAt(20); }
        if (upType==="lastmonth") { setLastMonth(recs); setPageLm(20); }
        if (upType==="inhouse")   setInhouse(recs);
        const now = new Date();
        const dateStr = `${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}-${now.getFullYear()}`;
        await supabase.from('tiktok_hub_settings').upsert({key:'last_imported',value:dateStr,updated_at:new Date().toISOString()});
        setLastImported(dateStr);
        setShowUp(false);
      };
      reader.readAsText(f);
    });
  };

  // ── card edit panel ──────────────────────────────────────────────────────────

  const openEdit = (r: VideoRow) => {
    setEditingId(r.id);
    setEditDraft({
      audioHook:    r.audioHook    || "",
      visualHook:   r.visualHook   || "",
      textHook:     r.textHook     || "",
      videoLength:  r.videoLength  || "",
      sellingPoints: pts(r.sellingPoints).join("\n"),
    });
  };

  const saveEdit = async (r: VideoRow) => {
    const sp = editDraft.sellingPoints.split("\n").map(s=>s.trim()).filter(Boolean).join(" | ");
    const fields: Override = { audioHook:editDraft.audioHook, visualHook:editDraft.visualHook, textHook:editDraft.textHook, videoLength:editDraft.videoLength, sellingPoints:sp };
    await supabase.from('tiktok_overrides').upsert({
      report_id:     r.id,
      audio_hook:    fields.audioHook,
      visual_hook:   fields.visualHook,
      text_hook:     fields.textHook,
      video_length:  fields.videoLength,
      selling_points:fields.sellingPoints,
      updated_at:    new Date().toISOString(),
    });
    setOverridesMap(prev => new Map(prev).set(r.id, fields));
    const upd = (rs: VideoRow[]) => rs.map(rec => rec.id===r.id ? {...rec,...fields} : rec);
    setAllTime(upd); setLastMonth(upd); setInhouse(upd);
    setEditingId(null);
  };

  const cancelEdit = () => setEditingId(null);
  const toggleTranscript = (id: string) => setTranscriptOpen(prev => { const s=new Set(prev); s.has(id)?s.delete(id):s.add(id); return s; });

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

  // ── sub-components ────────────────────────────────────────────────────────────

  const VideoCard = ({r, showFilter}: {r: VideoRow; showFilter: boolean}) => {
    const sellPts  = pts(r.sellingPoints).map(lbl);
    const tags     = (r.hashtags||"").split(" ").filter(Boolean);
    const hidden   = hiddenIds.has(r.videoId);
    const isEditing = editingId === r.id;
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
      <div style={{display:"flex",background:"#fff",borderRadius:14,overflow:"hidden",
        boxShadow:"0 1px 3px rgba(0,0,0,0.07)",
        border:`1px solid ${hidden&&adminMode?"#fca5a5":"#e5e7eb"}`,
        marginBottom:16, opacity:hidden&&adminMode?0.55:1}}>

        <div style={{flexShrink:0,width:325,background:"#0a0a0a"}}>
          {r.videoId ? (
            <div style={{width:325,height:578,overflow:"hidden",flexShrink:0}}>
              <iframe src={`https://www.tiktok.com/embed/v2/${r.videoId}`}
                style={{display:"block",width:325,height:738,border:"none"}}
                allowFullScreen allow="encrypted-media" loading="lazy" title={`@${r.creator}`}/>
            </div>
          ) : (
            <div style={{width:325,height:578,display:"flex",alignItems:"center",justifyContent:"center",color:"#444",flexDirection:"column",gap:6,fontSize:12}}>
              <span style={{fontSize:24}}>📹</span>No embed
            </div>
          )}
        </div>

        <div style={{flex:1,padding:"16px 20px",overflow:"hidden",minWidth:0,display:"flex",flexDirection:"column",gap:10}}>

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
                  ["visualHook",  "🎬 Visual Hook",  "textarea"],
                  ["textHook",    "🎣 Text Hook",    "textarea"],
                  ["videoLength", "⏱ Video Length",  "input"  ],
                ] as [keyof EditDraft, string, string][]
              ).map(([field, label, type]) => (
                <div key={field}>
                  <div style={{fontSize:10,color:"#6b7280",fontWeight:600,marginBottom:3}}>{label}</div>
                  {type==="input" ? (
                    <input value={editDraft[field]||""} onChange={e=>setEditDraft(d=>({...d,[field]:e.target.value}))}
                      style={{width:"100%",padding:"6px 10px",border:"1px solid #d1d5db",borderRadius:6,fontFamily:"inherit",fontSize:12,boxSizing:"border-box",outline:"none"}}/>
                  ) : (
                    <textarea value={editDraft[field]||""} onChange={e=>setEditDraft(d=>({...d,[field]:e.target.value}))} rows={2}
                      style={{display:"block",width:"100%",padding:"6px 10px",border:"1px solid #d1d5db",borderRadius:6,fontFamily:"inherit",fontSize:12,resize:"vertical",boxSizing:"border-box",outline:"none",lineHeight:1.5}}/>
                  )}
                </div>
              ))}
              <div>
                <div style={{fontSize:10,color:"#6b7280",fontWeight:600,marginBottom:3}}>✅ Selling Points <span style={{fontWeight:400,color:"#9ca3af"}}>— one per line</span></div>
                <textarea value={editDraft.sellingPoints||""} onChange={e=>setEditDraft(d=>({...d,sellingPoints:e.target.value}))} rows={3}
                  style={{display:"block",width:"100%",padding:"6px 10px",border:"1px solid #d1d5db",borderRadius:6,fontFamily:"inherit",fontSize:12,resize:"vertical",boxSizing:"border-box",outline:"none",lineHeight:1.6}}/>
              </div>
              <button onClick={()=>saveEdit(r)}
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
        <div style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)"}}>
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
            <div style={{display:"grid",gridTemplateColumns:`repeat(${Math.min(c.top3.length,3)},1fr)`,gap:16}}>
              {c.top3.map((v,i)=>(
                <div key={v.id||i} style={{minWidth:0}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                    <span style={{background:"#111",color:"#fff",borderRadius:6,padding:"3px 8px",fontSize:11,fontWeight:700}}>
                      {["1st","2nd","3rd"][i]||`#${i+1}`}
                    </span>
                    <span style={{fontWeight:700,color:"#16a34a",fontSize:13}}>{f$(v.revenue)}</span>
                    <span style={{fontSize:11,color:"#9ca3af"}}>{fN(v.itemsSold)} sold</span>
                    {v.datePosted && <span style={{fontSize:11,color:"#9ca3af"}}>· {v.datePosted}</span>}
                  </div>
                  {v.videoId ? (
                    <div style={{overflow:"hidden",borderRadius:10,height:Math.round(260*(16/9))}}>
                      <iframe src={`https://www.tiktok.com/embed/v2/${v.videoId}`}
                        style={{display:"block",width:"100%",height:738,border:"none"}}
                        allowFullScreen allow="encrypted-media" loading="lazy" title={`Top video ${i+1}`}/>
                    </div>
                  ) : (
                    <div style={{height:200,background:"#0a0a0a",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",color:"#555",fontSize:12}}>No embed</div>
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

  const slicedAt = visAllTime.slice(0, pageAt);
  const slicedLm = filteredLastMonth.slice(0, pageLm);
  const tabCount = {alltime:visAllTime.length, lastmonth:filteredLastMonth.length, inhouse:visInhouse.length, creators:filteredCreators.length};
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
    <div style={{fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif",background:"#f4f5f7",minHeight:"100vh"}}>

      <div style={{background:"#0c0c0c",padding:"16px 24px 0"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
          <span style={{fontSize:22}}>🐾</span>
          <div>
            <div style={{color:"#fff",fontWeight:800,fontSize:15,letterSpacing:"-0.02em"}}>RUFF LINERS</div>
            <div style={{color:"#444",fontSize:10,letterSpacing:"0.1em"}}>CREATOR HUB · TIKTOK SHOP</div>
          </div>
          <span style={{flex:1}}/>
          {lastImported && (
            <span style={{fontSize:11,color:"#666",padding:"4px 10px",borderRadius:20,background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",whiteSpace:"nowrap"}}>
              Date last Imported {lastImported}
            </span>
          )}
          {isAdmin && (
            <span style={{fontSize:11,fontWeight:700,padding:"4px 12px",borderRadius:20,letterSpacing:"0.05em",
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

        <div style={{display:"flex",overflowX:"auto"}}>
          {TABS.map(t=>{
            const active=tab===t.id;
            return (
              <button key={t.id} onClick={()=>setTab(t.id)}
                style={{padding:"12px 18px",border:"none",borderBottom:active?"3px solid #fff":"3px solid transparent",background:"none",cursor:"pointer",fontSize:13,fontWeight:active?700:400,color:active?"#fff":"#666",whiteSpace:"nowrap",fontFamily:"inherit",display:"flex",alignItems:"center",gap:7,transition:"color .12s"}}>
                {t.icon} {t.label}
                <span style={{background:active?"rgba(255,255,255,0.14)":"rgba(255,255,255,0.05)",color:active?"#fff":"#555",borderRadius:20,padding:"2px 8px",fontSize:10,fontWeight:700}}>
                  {tabCount[t.id as keyof typeof tabCount]}
                </span>
              </button>
            );
          })}
        </div>
      </div>

      {tab==="alltime" && (
        <>
          <div style={{background:"#fff",borderBottom:"1px solid #e5e7eb",padding:"10px 24px",display:"flex",gap:20,alignItems:"center",flexWrap:"wrap"}}>
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
          <div style={{background:"#fff",borderBottom:"1px solid #e5e7eb",padding:"10px 24px",display:"flex",gap:20,alignItems:"center",flexWrap:"wrap"}}>
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
          <div style={{background:"#fff",borderBottom:"1px solid #e5e7eb",padding:"10px 24px",display:"flex",gap:20,alignItems:"center",flexWrap:"wrap"}}>
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

      <div style={{maxWidth:1100,margin:"0 auto",padding:"20px"}}>

        {tab==="alltime" && (
          visAllTime.length===0
            ? <Empty msg='Click "Update Reports" above and upload the All-Time export from app.euka.ai/videos'/>
            : <>
                {slicedAt.map(r=><VideoCard key={r.id} r={r} showFilter={adminMode}/>)}
                {pageAt<visAllTime.length && (
                  <div style={{textAlign:"center",padding:"16px 0"}}>
                    <button onClick={()=>setPageAt(n=>n+20)}
                      style={{padding:"10px 28px",background:"#fff",border:"1px solid #d1d5db",borderRadius:10,cursor:"pointer",fontFamily:"inherit",fontSize:13,color:"#374151",fontWeight:500}}>
                      Show more ({visAllTime.length-pageAt} remaining)
                    </button>
                  </div>
                )}
              </>
        )}

        {tab==="lastmonth" && (
          filteredLastMonth.length===0
            ? <Empty msg='Click "Update Reports" above and upload the Last Month export from app.euka.ai/videos'/>
            : <>
                {slicedLm.map(r=><VideoCard key={r.id} r={r} showFilter={false}/>)}
                {pageLm<filteredLastMonth.length && (
                  <div style={{textAlign:"center",padding:"16px 0"}}>
                    <button onClick={()=>setPageLm(n=>n+20)}
                      style={{padding:"10px 28px",background:"#fff",border:"1px solid #d1d5db",borderRadius:10,cursor:"pointer",fontFamily:"inherit",fontSize:13,color:"#374151",fontWeight:500}}>
                      Show more ({filteredLastMonth.length-pageLm} remaining)
                    </button>
                  </div>
                )}
              </>
        )}

        {tab==="inhouse" && (
          visInhouse.length===0
            ? <Empty msg="In-House Content report source TBD — upload when available"/>
            : visInhouse.map(r=><VideoCard key={r.id} r={r} showFilter={false}/>)
        )}

        {tab==="creators" && (
          filteredCreators.length===0
            ? <Empty msg='Upload the All-Time report first — Top Creators are computed from that data'/>
            : filteredCreators.map((c,i)=><CreatorCard key={c.creator} c={c} idx={i}/>)
        )}
      </div>


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
    </div>
  );
}
