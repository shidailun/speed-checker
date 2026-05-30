import React, { useState, useRef, useEffect } from 'react';
import {
  View,
  Text,
  TextInput,
  TouchableOpacity,
  Modal,
  ScrollView,
  StyleSheet,
  ActivityIndicator,
  KeyboardAvoidingView,
  PanResponder,
  Platform,
  Alert,
} from 'react-native';
import * as Updates from 'expo-updates';
import * as DocumentPicker from 'expo-document-picker';
import * as LegacyFS from 'expo-file-system/legacy';
import * as Sharing from 'expo-sharing';
import { Audio } from 'expo-av';
import * as XLSX from 'xlsx';
import { StatusBar } from 'expo-status-bar';

const C = {
  bg:      '#0d1117',
  surface: '#161b22',
  accent:  '#4fc3f7',
  green:   '#4CAF50',
  red:     '#ef5350',
  text:    '#e6edf3',
  muted:   '#8b949e',
};

type Entry = { text: string; filename: string };

interface AppConfig {
  xlsxLocalPath?:   string;
  xlsxName?:        string;
  audioFolderUri?:  string;
  lastSheet?:       string;
  lastIdx?:         number;
}

const CONFIG_PATH = LegacyFS.documentDirectory
  ? LegacyFS.documentDirectory + 'app_config.json'
  : null;

async function readConfig(): Promise<AppConfig> {
  if (!CONFIG_PATH) return {};
  try {
    const info = await LegacyFS.getInfoAsync(CONFIG_PATH);
    if (!info.exists) return {};
    return JSON.parse(await LegacyFS.readAsStringAsync(CONFIG_PATH)) as AppConfig;
  } catch { return {}; }
}

async function writeConfig(cfg: AppConfig): Promise<void> {
  if (!CONFIG_PATH) return;
  try { await LegacyFS.writeAsStringAsync(CONFIG_PATH, JSON.stringify(cfg)); } catch {}
}

async function fetchWorkbook(uri: string): Promise<XLSX.WorkBook> {
  const resp = await fetch(uri);
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  return XLSX.read(new Uint8Array(await resp.arrayBuffer()), { type: 'array' });
}

function workingName(xlsxName: string): string {
  return xlsxName.replace(/\.(xlsx|xls)$/i, '_working.xlsx');
}

function matchesPattern(text: string, pattern: string): boolean {
  const p = pattern.trim();
  if (!p) return true;
  const leadDash  = p.startsWith('-');
  const trailDash = p.endsWith('-');
  const sound     = p.replace(/^-|-$/g, '').toLowerCase();
  if (!sound) return true;
  const words = text.toLowerCase().split(/\s+/).filter(Boolean);
  if (leadDash && trailDash) {
    return words.some(w => { const i = w.indexOf(sound); return i > 0 && i + sound.length < w.length; });
  } else if (leadDash) {
    return words.some(w => w.endsWith(sound));
  } else if (trailDash) {
    return words.some(w => w.startsWith(sound));
  } else {
    return words.some(w => w.includes(sound));
  }
}

function applyReplacement(text: string, pattern: string, replacement: string): string {
  const p = pattern.trim();
  if (!p) return text;
  const leadDash  = p.startsWith('-');
  const trailDash = p.endsWith('-');
  const sound     = p.replace(/^-|-$/g, '').toLowerCase();
  if (!sound) return text;
  const tokens = text.split(/(\s+)/);
  return tokens.map(token => {
    if (/^\s+$/.test(token)) return token;
    const w = token.toLowerCase();
    if (leadDash && trailDash) {
      const i = w.indexOf(sound);
      if (i > 0 && i + sound.length < w.length)
        return token.slice(0, i) + replacement + token.slice(i + sound.length);
    } else if (leadDash) {
      if (w.endsWith(sound))
        return token.slice(0, token.length - sound.length) + replacement;
    } else if (trailDash) {
      if (w.startsWith(sound))
        return replacement + token.slice(sound.length);
    } else {
      return token.replace(new RegExp(sound.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), replacement);
    }
    return token;
  }).join('');
}

function computeFilteredIndices(allEntries: Entry[], filter: string): number[] {
  if (!filter.trim()) return allEntries.map((_, i) => i);
  const patterns = filter.split(',').map(p => p.trim()).filter(Boolean);
  return allEntries.reduce<number[]>((acc, entry, i) => {
    if (patterns.some(p => matchesPattern(entry.text, p))) acc.push(i);
    return acc;
  }, []);
}

export default function App() {
  const [entries,       setEntries]       = useState<Entry[]>([]);
  const [idx,           setIdx]           = useState(0);
  const [speed,         setSpeed]         = useState(1.0);
  const [editText,      setEditText]      = useState('');
  const [playing,       setPlaying]       = useState(false);
  const [status,        setStatus]        = useState('Loading…');
  const [loading,       setLoading]       = useState(false);
  const [sheets,        setSheets]        = useState<string[]>([]);
  const [curSheet,      setCurSheet]      = useState<string | null>(null);
  const [sheetModal,    setSheetModal]    = useState(false);
  const [jumpText,      setJumpText]      = useState('');
  const [audioCount,    setAudioCount]    = useState(0);
  const [xlsxLabel,     setXlsxLabel]     = useState('');
  const [position,      setPosition]      = useState(0);
  const [duration,      setDuration]      = useState(0);
  const [filenameModal, setFilenameModal] = useState(false);
  const [editFilename,  setEditFilename]  = useState('');
  const [filterInput,   setFilterInput]   = useState('');
  const [replaceInput,  setReplaceInput]  = useState('');
  const [displayIdx,    setDisplayIdx]    = useState(0);
  const [filteredCount, setFilteredCount] = useState(0);
  const [aboutModal,    setAboutModal]    = useState(false);
  const [logLines,      setLogLines]      = useState<string[]>([]);
  const [cutMode,       setCutMode]       = useState(false);
  const [cutIn,         setCutIn]         = useState<number | null>(null);
  const [cutOut,        setCutOut]        = useState<number | null>(null);
  const [cutInText,     setCutInText]     = useState('');
  const [cutOutText,    setCutOutText]    = useState('');
  const [playingPreview, setPlayingPreview] = useState(false);

  const soundRef             = useRef<Audio.Sound | null>(null);
  const workbookRef          = useRef<XLSX.WorkBook | null>(null);
  const audioMapRef          = useRef<Record<string, string>>({});
  const entriesRef           = useRef<Entry[]>([]);
  const allEntriesRef        = useRef<Entry[]>([]);
  const filteredIndicesRef   = useRef<number[]>([]);
  const displayIdxRef        = useRef(0);
  const idxRef               = useRef(0);
  const speedRef             = useRef(1.0);
  const xlsxLocalRef         = useRef<string | null>(null);
  const xlsxRemoteRef        = useRef<string | null>(null);
  const curSheetRef          = useRef<string | null>(null);
  const configRef            = useRef<AppConfig>({});
  const trackWidthRef        = useRef(1);
  const saveTimerRef         = useRef<ReturnType<typeof setTimeout> | null>(null);
  const speedHoldTimeoutRef  = useRef<ReturnType<typeof setTimeout> | null>(null);
  const speedHoldIntervalRef = useRef<ReturnType<typeof setInterval> | null>(null);
  const undoRef              = useRef<{ globalIdx: number; oldText: string }[] | null>(null);
  const cutPreviewRef        = useRef(false);
  const cutInRef             = useRef<number | null>(null);
  const cutOutRef            = useRef<number | null>(null);
  const cutUndoRef           = useRef<{ globalIdx: number; oldFilename: string } | null>(null);

  const navRef = useRef({ goNext: () => {}, goPrev: () => {} });
  const swipeResponder = useRef(
    PanResponder.create({
      onStartShouldSetPanResponder: () => false,
      onMoveShouldSetPanResponder:  (_, gs) =>
        Math.abs(gs.dx) > Math.abs(gs.dy) * 1.5 && Math.abs(gs.dx) > 12,
      onPanResponderRelease: (_, gs) => {
        if (Math.abs(gs.dx) < 50) return;
        if (gs.dx < 0) navRef.current.goNext();
        else           navRef.current.goPrev();
      },
    })
  ).current;

  useEffect(() => {
    if (!__DEV__)
      Updates.checkForUpdateAsync()
        .then(({ isAvailable }) => { if (isAvailable) return Updates.fetchUpdateAsync(); })
        .then(r => { if (r) Updates.reloadAsync(); })
        .catch(() => {});
  }, []);

  useEffect(() => {
    Audio.setAudioModeAsync({
      allowsRecordingIOS:      false,
      staysActiveInBackground: false,
      shouldDuckAndroid:       true,
    });

    (async () => {
      const cfg = await readConfig();
      configRef.current = cfg;

      let audioMap: Record<string, string> = {};
      if (cfg.audioFolderUri && Platform.OS !== 'web') {
        try {
          const uris = await LegacyFS.StorageAccessFramework.readDirectoryAsync(cfg.audioFolderUri);
          audioMap = buildAudioMap(uris);
          audioMapRef.current = audioMap;
          setAudioCount(Object.keys(audioMap).length);
        } catch {}
      }

      let excelOk = false;
      if (cfg.xlsxLocalPath && Platform.OS !== 'web') {
        try {
          const info = await LegacyFS.getInfoAsync(cfg.xlsxLocalPath);
          if (info.exists) {
            setStatus(`Loading ${cfg.xlsxName ?? 'last file'}…`);
            const wb = await fetchWorkbook(cfg.xlsxLocalPath);
            workbookRef.current = wb;
            xlsxLocalRef.current = cfg.xlsxLocalPath;
            setXlsxLabel(cfg.xlsxName ?? '');
            const detected = detectSheets(wb);
            setSheets(detected);
            if (detected.length === 1) {
              loadSheet(detected[0], wb, audioMap, cfg.lastIdx ?? 0);
            } else if (detected.length > 1) {
              const resume = cfg.lastSheet && detected.includes(cfg.lastSheet) ? cfg.lastSheet : null;
              if (resume) {
                loadSheet(resume, wb, audioMap, cfg.lastIdx ?? 0);
              } else {
                setStatus(`${cfg.xlsxName} — ${detected.length} sheets. Tap "Sheet".`);
                setSheetModal(true);
              }
            } else {
              setStatus('No matching sheets found. Tap "Excel" to pick another file.');
            }
            excelOk = true;
          }
        } catch {
          setStatus('Could not load last file. Tap "Excel" to reload.');
        }
      }

      if (!excelOk && Object.keys(audioMap).length === 0) {
        setStatus('Tap "Excel" then "Audio" to begin.');
      } else if (!excelOk) {
        setStatus(`${Object.keys(audioMap).length} audio files ready. Tap "Excel" to begin.`);
      }
    })();

    return () => { soundRef.current?.unloadAsync().catch(() => {}); };
  }, []);

  useEffect(() => {
    if (!workbookRef.current || !curSheetRef.current || !entriesRef.current.length) return;
    const i   = idxRef.current;
    const cur = entriesRef.current[i];
    if (!cur || editText === cur.text) return;
    const capturedIdx  = i;
    const capturedText = editText;
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => persistEntry(capturedIdx, capturedText), 1500);
    return () => { if (saveTimerRef.current) clearTimeout(saveTimerRef.current); };
  }, [editText]);

  useEffect(() => { cutInRef.current  = cutIn;  }, [cutIn]);
  useEffect(() => { cutOutRef.current = cutOut; }, [cutOut]);

  // During preview: skip over the cut region mid-playback
  useEffect(() => {
    if (!cutPreviewRef.current || cutIn === null || cutOut === null) return;
    if (position >= cutIn && position < cutOut)
      soundRef.current?.setPositionAsync(cutOut).catch(() => {});
  }, [position, cutIn, cutOut]);

  function buildAudioMap(uris: string[]): Record<string, string> {
    const ext = /\.(wav|mp3|flac|m4a|ogg|aac)$/i;
    const map: Record<string, string> = {};
    for (const uri of uris) {
      const name = extractFilename(uri).toLowerCase();
      if (ext.test(name)) map[name] = uri;
    }
    return map;
  }

  function buildAudioMapFromAssets(assets: DocumentPicker.DocumentPickerAsset[]): Record<string, string> {
    const ext = /\.(wav|mp3|flac|m4a|ogg|aac)$/i;
    const map: Record<string, string> = {};
    for (const asset of assets) {
      const name = (asset.name ?? extractFilename(asset.uri)).toLowerCase();
      if (ext.test(name)) map[name] = asset.uri;
    }
    return map;
  }

  function extractFilename(uri: string): string {
    try { return decodeURIComponent(uri).split('/').pop() ?? ''; }
    catch { return uri.split('/').pop() ?? ''; }
  }

  function extractFolderName(uri: string): string {
    try {
      const decoded = decodeURIComponent(uri);
      // SAF URIs look like "content://.../.../primary:Path/to/folder"
      const afterColon = decoded.slice(decoded.lastIndexOf(':') + 1);
      const parts = afterColon.split('/').filter(Boolean);
      return parts[parts.length - 1] ?? '';
    } catch { return ''; }
  }

  function findAudioUri(filename: string, map: Record<string, string>): string | null {
    const lower = filename.toLowerCase();
    if (map[lower]) return map[lower];
    const stem = lower.replace(/\.[^.]+$/, '');
    return map[`${stem}.wav`] ?? map[`${stem}.mp3`] ?? map[`${stem}.flac`] ?? null;
  }

  function detectSheets(wb: XLSX.WorkBook): string[] {
    return wb.SheetNames.filter((name) => {
      const lower = name.toLowerCase();
      if (!(lower.includes('sentence') || lower.includes('word'))) return false;
      if (!name.includes("'")) return false;
      const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[name], { header: 1 });
      return rows.some((r) => Array.isArray(r) && r.length >= 2 && r[1]);
    });
  }

  function parseEntries(wb: XLSX.WorkBook, sheetName: string): Entry[] {
    const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[sheetName], { header: 1 });
    const result: Entry[] = [];
    for (const row of rows) {
      if (!Array.isArray(row) || row.length < 2) continue;
      const text = String(row[0] ?? '').trim(), filename = String(row[1] ?? '').trim();
      if (text && filename) result.push({ text, filename });
    }
    return result;
  }

  const stopSound = async () => {
    if (soundRef.current) {
      try { await soundRef.current.stopAsync(); }   catch {}
      try { await soundRef.current.unloadAsync(); } catch {}
      soundRef.current = null;
    }
    setPlaying(false);
    setPosition(0);
    setDuration(0);
    setPlayingPreview(false);
  };

  const playUri = async (uri: string) => {
    await stopSound();
    try {
      const { sound } = await Audio.Sound.createAsync(
        { uri },
        { shouldPlay: true, rate: speedRef.current, shouldCorrectPitch: false, volume: 1.0, progressUpdateIntervalMillis: 50 },
        (s) => {
          if (!s.isLoaded) return;
          if (s.didJustFinish) setPlaying(false);
          setPosition(s.positionMillis ?? 0);
          if (s.durationMillis) setDuration(s.durationMillis);
        },
      );
      soundRef.current = sound;
      setPlaying(true);
    } catch (e) { setStatus(`Audio error: ${String(e)}`); }
  };

  const playCurrentEntry = async (list?: Entry[], i?: number, map?: Record<string, string>, preview = false) => {
    cutPreviewRef.current = preview;
    const L = list ?? entriesRef.current;
    const n = i    ?? idxRef.current;
    const m = map  ?? audioMapRef.current;
    if (!L.length || n >= L.length) return;
    const uri = findAudioUri(L[n].filename, m);
    if (!uri) { setStatus(`Audio not found: ${L[n].filename}`); return; }
    setStatus(L[n].filename);
    await playUri(uri);
  };

  const startSpeedHold = (delta: number) => {
    const next = Math.round(Math.max(0.5, Math.min(1.5, speedRef.current + delta)) * 100) / 100;
    speedRef.current = next; setSpeed(next);
    if (soundRef.current) soundRef.current.setRateAsync(next, false).catch(() => {});
    speedHoldTimeoutRef.current = setTimeout(() => {
      speedHoldIntervalRef.current = setInterval(() => {
        const n = Math.round(Math.max(0.5, Math.min(1.5, speedRef.current + delta)) * 100) / 100;
        speedRef.current = n; setSpeed(n);
        if (soundRef.current) soundRef.current.setRateAsync(n, false).catch(() => {});
      }, 120);
    }, 400);
  };

  const stopSpeedHold = () => {
    if (speedHoldTimeoutRef.current)  { clearTimeout(speedHoldTimeoutRef.current);   speedHoldTimeoutRef.current  = null; }
    if (speedHoldIntervalRef.current) { clearInterval(speedHoldIntervalRef.current); speedHoldIntervalRef.current = null; }
  };

  const pickExcel = async () => {
    try {
      const result = await DocumentPicker.getDocumentAsync({
        type: ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
               'application/vnd.ms-excel', '*/*'],
        copyToCacheDirectory: false,
      });
      if (result.canceled) return;
      const asset     = result.assets[0];
      const fileUri   = asset.uri;
      const fileName  = asset.name ?? extractFilename(fileUri);
      setStatus(`Reading ${fileName}…`);
      const wb = await fetchWorkbook(fileUri);
      workbookRef.current = wb;
      xlsxRemoteRef.current = fileUri;
      setXlsxLabel(fileName);

      if (Platform.OS !== 'web') {
        const localPath = LegacyFS.documentDirectory + workingName(fileName);
        const base64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx', compression: true });
        await LegacyFS.writeAsStringAsync(localPath, base64, { encoding: LegacyFS.EncodingType.Base64 });
        xlsxLocalRef.current = localPath;
        const cfg: AppConfig = { ...configRef.current, xlsxLocalPath: localPath, xlsxName: fileName };
        configRef.current = cfg; writeConfig(cfg);
      } else {
        xlsxLocalRef.current = fileUri;
        configRef.current = { ...configRef.current, xlsxName: fileName };
      }

      const detected = detectSheets(wb);
      setSheets(detected);
      if (detected.length === 0) {
        setStatus("No matching sheets. Names must contain 'sentence' or 'word' plus apostrophe.");
      } else if (detected.length === 1) {
        loadSheet(detected[0], wb, audioMapRef.current);
      } else {
        setStatus(`${fileName} loaded — pick an audio folder to auto-select the sheet.`);
      }
    } catch (e) { setStatus(`Excel error: ${String(e)}`); }
  };

  const pickAudioFolder = async () => {
    if (Platform.OS === 'web') {
      try {
        const result = await DocumentPicker.getDocumentAsync({
          type: ['audio/*'],
          multiple: true,
          copyToCacheDirectory: false,
        });
        if (result.canceled) return;
        const map   = buildAudioMapFromAssets(result.assets);
        audioMapRef.current = map;
        const count = Object.keys(map).length;
        setAudioCount(count);
        setStatus(`✓ ${count} audio files loaded.`);
        if (entriesRef.current.length > 0)
          await playCurrentEntry(entriesRef.current, idxRef.current, map);
      } catch (e) { setStatus(`Audio error: ${String(e)}`); }
      return;
    }

    try {
      const perm = await LegacyFS.StorageAccessFramework.requestDirectoryPermissionsAsync();
      if (!perm.granted) return;
      setStatus('Scanning folder…');
      const uris  = await LegacyFS.StorageAccessFramework.readDirectoryAsync(perm.directoryUri);
      const map   = buildAudioMap(uris);
      audioMapRef.current = map;
      const count = Object.keys(map).length;
      setAudioCount(count);
      const cfg = { ...configRef.current, audioFolderUri: perm.directoryUri };
      configRef.current = cfg; writeConfig(cfg);

      const folderName = extractFolderName(perm.directoryUri).toLowerCase();
      const wb = workbookRef.current;
      let autoSwitched = false;
      if (wb) {
        const detected = detectSheets(wb);
        const matchingSheet = detected.find(sh => {
          const shLower = sh.toLowerCase();
          return (folderName.includes('word') && shLower.includes('word')) ||
                 (folderName.includes('sentence') && shLower.includes('sentence'));
        });
        if (matchingSheet && matchingSheet !== curSheetRef.current) {
          loadSheet(matchingSheet, wb, map);
          autoSwitched = true;
        }
      }

      if (!autoSwitched) {
        setStatus(`✓ ${count} audio files found.`);
        if (entriesRef.current.length > 0)
          await playCurrentEntry(entriesRef.current, idxRef.current, map);
      }
    } catch (e) { setStatus(`Folder error: ${String(e)}`); }
  };

  const doShare = async () => {
    const wb        = workbookRef.current;
    const localPath = xlsxLocalRef.current;
    if (!wb || !localPath) { setStatus('Nothing to share yet.'); return; }
    const xlsxName  = configRef.current.xlsxName ?? 'export.xlsx';
    const trimUris  = Object.entries(audioMapRef.current)
      .filter(([name]) => name.endsWith('_trim.wav'));

    if (Platform.OS === 'web') {
      try {
        const array = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
        const blob  = new Blob([array], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url   = URL.createObjectURL(blob);
        const link  = document.createElement('a');
        link.href = url; link.download = xlsxName; link.click();
        URL.revokeObjectURL(url);
        setStatus('Downloaded ✓');
      } catch (e) { setStatus(`Download error: ${String(e)}`); }
      return;
    }

    setLoading(true);
    try {
      // Always save Excel to local path first
      const xlsxB64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx', compression: true });
      await LegacyFS.writeAsStringAsync(localPath, xlsxB64, { encoding: LegacyFS.EncodingType.Base64 });

      if (trimUris.length === 0) {
        // No trim files — share Excel only (original behaviour)
        await Sharing.shareAsync(localPath, {
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          dialogTitle: 'Share spreadsheet',
          UTI: 'com.microsoft.excel.xlsx',
        });
        setStatus('Shared ✓');
        return;
      }

      // Build zip: Excel + all _trim.wav files
      setStatus(`Building zip (${trimUris.length} trim file${trimUris.length > 1 ? 's' : ''})…`);
      const xlsxBytes = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
      const zipFiles: { name: string; data: Uint8Array }[] = [
        { name: xlsxName, data: new Uint8Array(xlsxBytes) },
      ];
      for (const [name, uri] of trimUris) {
        try {
          const resp = await fetch(uri);
          const buf  = await resp.arrayBuffer();
          zipFiles.push({ name, data: new Uint8Array(buf) });
        } catch { /* skip unreadable files */ }
      }
      const zipData  = buildZip(zipFiles);
      const zipName  = xlsxName.replace(/\.[^.]+$/, '') + '_export.zip';
      const zipPath  = (LegacyFS.documentDirectory ?? '') + zipName;
      const u8 = zipData;
      let bin = ''; const ch = 8192;
      for (let j = 0; j < u8.length; j += ch)
        bin += String.fromCharCode(...u8.subarray(j, Math.min(j + ch, u8.length)));
      await LegacyFS.writeAsStringAsync(zipPath, btoa(bin), { encoding: LegacyFS.EncodingType.Base64 });
      await Sharing.shareAsync(zipPath, { mimeType: 'application/zip', dialogTitle: 'Export transcriptions' });
      setStatus(`Shared zip (${zipFiles.length} files) ✓`);
    } catch (e) { setStatus(`Share error: ${String(e)}`); }
    finally { setLoading(false); }
  };

  const persistEntry = async (i: number, text: string) => {
    const wb    = workbookRef.current;
    const sheet = curSheetRef.current;
    const list  = entriesRef.current;
    if (!wb || !sheet || !list.length || i >= list.length || !text.trim()) return;
    const ws   = wb.Sheets[sheet];
    const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 });
    let dataIdx = 0;
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r] as unknown[];
      if (Array.isArray(row) && row.length >= 2 && row[0] && row[1]) {
        if (dataIdx === i) {
          ws[XLSX.utils.encode_cell({ r, c: 0 })] = { t: 's', v: text.trim(), w: text.trim() };
          break;
        }
        dataIdx++;
      }
    }
    const oldText = list[i].text;
    entriesRef.current[i] = { ...list[i], text: text.trim() };
    setEntries(prev => { const n = [...prev]; if (n[i]) n[i] = { ...n[i], text: text.trim() }; return n; });
    writeConfig({ ...configRef.current, lastIdx: i });
    const lp = xlsxLocalRef.current;
    if (!lp || Platform.OS === 'web') return;
    const base64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx', compression: true });
    try {
      await LegacyFS.writeAsStringAsync(lp, base64, { encoding: LegacyFS.EncodingType.Base64 });
      setStatus(`Auto-saved ✓  ${list[i].filename}`);
      const diff = wordDiff(oldText, text.trim());
      if (diff) addLog(`${list[i].filename}  ${diff}`);
    } catch (e) { setStatus(`Auto-save error: ${String(e)}`); }
  };

  const persistFilename = async (i: number, newFilename: string) => {
    const wb    = workbookRef.current;
    const sheet = curSheetRef.current;
    const list  = entriesRef.current;
    if (!wb || !sheet || !list.length || i >= list.length || !newFilename.trim()) return;
    setFilenameModal(false);
    const ws   = wb.Sheets[sheet];
    const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 });
    let dataIdx = 0;
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r] as unknown[];
      if (Array.isArray(row) && row.length >= 2 && row[0] && row[1]) {
        if (dataIdx === i) {
          ws[XLSX.utils.encode_cell({ r, c: 1 })] = { t: 's', v: newFilename.trim(), w: newFilename.trim() };
          break;
        }
        dataIdx++;
      }
    }
    entriesRef.current[i] = { ...list[i], filename: newFilename.trim() };
    setEntries(prev => { const n = [...prev]; if (n[i]) n[i] = { ...n[i], filename: newFilename.trim() }; return n; });
    const lp = xlsxLocalRef.current;
    if (!lp || Platform.OS === 'web') return;
    const base64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx', compression: true });
    try {
      await LegacyFS.writeAsStringAsync(lp, base64, { encoding: LegacyFS.EncodingType.Base64 });
      setStatus(`Filename updated ✓  ${newFilename.trim()}`);
    } catch (e) { setStatus(`Save error: ${String(e)}`); }
  };

  const deleteEntry = async (i: number) => {
    const entry = entriesRef.current[i];
    if (!entry) return;
    Alert.alert(
      'Delete Entry',
      `Remove "${entry.filename}" from the spreadsheet?\n\nThe audio file is NOT deleted.`,
      [
        { text: 'Cancel', style: 'cancel' },
        { text: 'Delete', style: 'destructive', onPress: async () => {
          setFilenameModal(false);
          const wb    = workbookRef.current;
          const sheet = curSheetRef.current;
          if (!wb || !sheet) return;
          const ws   = wb.Sheets[sheet];
          const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 });
          let di = 0, removeRow = -1;
          for (let r = 0; r < rows.length; r++) {
            const row = rows[r] as unknown[];
            if (Array.isArray(row) && row.length >= 2 && row[0] && row[1]) {
              if (di === i) { removeRow = r; break; }
              di++;
            }
          }
          if (removeRow >= 0) removeXlsxRow(ws, removeRow);
          const all = allEntriesRef.current;
          allEntriesRef.current = [...all.slice(0, i), ...all.slice(i + 1)];
          entriesRef.current    = allEntriesRef.current;
          const fi = filteredIndicesRef.current.filter(x => x !== i).map(x => x > i ? x - 1 : x);
          filteredIndicesRef.current = fi;
          setEntries([...entriesRef.current]);
          setFilteredCount(fi.length);
          addLog(`Deleted entry: ${entry.filename}`);
          const lp = xlsxLocalRef.current;
          if (lp && Platform.OS !== 'web') {
            const base64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx', compression: true });
            try { await LegacyFS.writeAsStringAsync(lp, base64, { encoding: LegacyFS.EncodingType.Base64 }); }
            catch (e) { setStatus(`Save error: ${String(e)}`); return; }
          }
          if (fi.length > 0) {
            const nextDi = Math.min(displayIdxRef.current, fi.length - 1);
            await gotoEntry(fi[nextDi]);
          } else {
            setEditText(''); setStatus('No entries remaining.');
          }
        }},
      ],
    );
  };

  const loadSheet = (
    sheetName: string,
    wb?:       XLSX.WorkBook,
    map?:      Record<string, string>,
    startIdx?: number,
  ) => {
    const workbook = wb ?? workbookRef.current;
    if (!workbook) return;
    setSheetModal(false);
    curSheetRef.current = sheetName; setCurSheet(sheetName);
    const cfg = { ...configRef.current, lastSheet: sheetName };
    configRef.current = cfg; writeConfig(cfg);
    const loaded = parseEntries(workbook, sheetName);
    allEntriesRef.current      = loaded;
    entriesRef.current         = loaded; setEntries(loaded);
    filteredIndicesRef.current = loaded.map((_, i) => i);
    setFilteredCount(loaded.length);
    setFilterInput('');
    const at = Math.min(startIdx ?? 0, Math.max(0, loaded.length - 1));
    idxRef.current = at; setIdx(at);
    displayIdxRef.current = at; setDisplayIdx(at);
    setStatus(`'${sheetName}' — ${loaded.length} entries.${at > 0 ? ` Resuming #${at + 1}.` : ''}`);
    if (loaded.length > 0) {
      setEditText(loaded[at].text);
      const m = map ?? audioMapRef.current;
      if (Object.keys(m).length > 0) playCurrentEntry(loaded, at, m);
    }
  };

  const gotoEntry = async (n: number) => {
    await stopSound();
    idxRef.current = n; setIdx(n);
    const di = filteredIndicesRef.current.indexOf(n);
    if (di >= 0) { displayIdxRef.current = di; setDisplayIdx(di); }
    writeConfig({ ...configRef.current, lastIdx: n });
    const entry = entriesRef.current[n];
    if (!entry) return;
    setEditText(entry.text);
    const uri = findAudioUri(entry.filename, audioMapRef.current);
    if (uri) { setStatus(entry.filename); await playUri(uri); }
    else       setStatus(`Audio not found: ${entry.filename}`);
  };

  const goNext = async () => {
    const fi     = filteredIndicesRef.current;
    const nextDi = displayIdxRef.current + 1;
    if (nextDi >= fi.length) { setStatus('Last entry.'); return; }
    await gotoEntry(fi[nextDi]);
  };

  const goPrev = async () => {
    const fi     = filteredIndicesRef.current;
    const prevDi = displayIdxRef.current - 1;
    if (prevDi < 0) return;
    await gotoEntry(fi[prevDi]);
  };

  navRef.current.goNext = goNext;
  navRef.current.goPrev = goPrev;

  const doJump = async () => {
    const n = parseInt(jumpText, 10);
    if (isNaN(n)) return;
    setJumpText('');
    const fi = filteredIndicesRef.current;
    const di = Math.max(0, Math.min(n - 1, fi.length - 1));
    await gotoEntry(fi[di]);
  };

  const applyFilter = () => {
    const all = allEntriesRef.current;
    if (!all.length) return;
    const fi = computeFilteredIndices(all, filterInput);
    filteredIndicesRef.current = fi;
    setFilteredCount(fi.length);
    if (fi.length > 0) {
      displayIdxRef.current = 0; setDisplayIdx(0);
      gotoEntry(fi[0]);
    } else {
      setStatus('No matches found.');
    }
  };

  const addLog = (msg: string) =>
    setLogLines(prev => [`${logTime()}  ${msg}`, ...prev].slice(0, 50));

  const playBottomBand = async () => {
    const canCut = cutIn !== null && cutOut !== null && cutIn < cutOut!;
    if (canCut) {
      await doPreview();
    } else {
      // Play original on bottom band for marking
      const entry = entriesRef.current[idxRef.current];
      if (!entry) return;
      const uri = findAudioUri(entry.filename, audioMapRef.current);
      if (!uri) { setStatus(`Audio not found: ${entry.filename}`); return; }
      await stopSound();
      setStatus(entry.filename);
      setPlayingPreview(true);
      await playUri(uri);
    }
  };

  const doPreview = async () => {
    const entry = entriesRef.current[idxRef.current];
    if (!entry || cutIn === null || cutOut === null || cutIn >= cutOut) {
      setStatus('Set In and Out points first.'); return;
    }
    const uri = findAudioUri(entry.filename, audioMapRef.current);
    if (!uri) { setStatus('Audio not found.'); return; }
    setLoading(true);
    try {
      const buffer = await (await fetch(uri)).arrayBuffer();
      if (Platform.OS === 'web') {
        const AudioCtx = (window as any).AudioContext || (window as any).webkitAudioContext;
        const ctx = new AudioCtx() as AudioContext;
        const dec = await ctx.decodeAudioData(buffer.slice(0));
        const sr = dec.sampleRate, inS = Math.floor(cutIn / 1000 * sr), outS = Math.floor(cutOut / 1000 * sr);
        const nb = ctx.createBuffer(dec.numberOfChannels, dec.length - (outS - inS), sr);
        for (let ch = 0; ch < dec.numberOfChannels; ch++) {
          const s = dec.getChannelData(ch), d = nb.getChannelData(ch);
          d.set(s.subarray(0, inS), 0); d.set(s.subarray(outS), inS);
        }
        const blob = new Blob([audioBufferToWav(nb)], { type: 'audio/wav' });
        const url  = URL.createObjectURL(blob);
        setPlayingPreview(true);
        await playUri(url);
        setTimeout(() => URL.revokeObjectURL(url), 60000);
      } else {
        if (!entry.filename.toLowerCase().endsWith('.wav')) {
          setStatus('Preview only supports WAV on Android.'); return;
        }
        const newBuf = cutWavBuffer(buffer, cutIn, cutOut);
        const u8 = new Uint8Array(newBuf);
        let bin = ''; const ch = 8192;
        for (let j = 0; j < u8.length; j += ch)
          bin += String.fromCharCode(...u8.subarray(j, Math.min(j + ch, u8.length)));
        const tmp = LegacyFS.documentDirectory + '_preview_cut.wav';
        await LegacyFS.writeAsStringAsync(tmp, btoa(bin), { encoding: LegacyFS.EncodingType.Base64 });
        setPlayingPreview(true);
        await playUri(tmp);
      }
    } catch (e) { setStatus(`Preview error: ${String(e)}`); }
    finally { setLoading(false); }
  };

  const markIn = async () => {
    const st  = soundRef.current ? await soundRef.current.getStatusAsync() : null;
    const pos = st && st.isLoaded ? st.positionMillis : position;
    setCutIn(pos); setCutInText((pos / 1000).toFixed(3));
  };
  const markOut = async () => {
    const st  = soundRef.current ? await soundRef.current.getStatusAsync() : null;
    const pos = st && st.isLoaded ? st.positionMillis : position;
    setCutOut(pos); setCutOutText((pos / 1000).toFixed(3));
  };

  const onCutInTextChange = (t: string) => {
    setCutInText(t);
    const ms = parseFloat(t.replace(',', '.')) * 1000;
    if (!isNaN(ms) && ms >= 0) setCutIn(ms);
  };
  const onCutOutTextChange = (t: string) => {
    setCutOutText(t);
    const ms = parseFloat(t.replace(',', '.')) * 1000;
    if (!isNaN(ms) && ms >= 0) setCutOut(ms);
  };

  const nudgeHoldTimeout  = useRef<ReturnType<typeof setTimeout>  | null>(null);
  const nudgeHoldInterval = useRef<ReturnType<typeof setInterval> | null>(null);

  const applyNudge = (which: 'in' | 'out', delta: number) => {
    if (which === 'in') {
      setCutIn(prev => {
        const next = Math.max(0, (prev ?? 0) + delta);
        setCutInText((next / 1000).toFixed(3));
        return next;
      });
    } else {
      setCutOut(prev => {
        const next = Math.max(0, (prev ?? 0) + delta);
        setCutOutText((next / 1000).toFixed(3));
        return next;
      });
    }
  };

  const startNudgeHold = (which: 'in' | 'out', delta: number) => {
    applyNudge(which, delta);
    nudgeHoldTimeout.current = setTimeout(() => {
      nudgeHoldInterval.current = setInterval(() => applyNudge(which, delta), 80);
    }, 350);
  };
  const stopNudgeHold = () => {
    if (nudgeHoldTimeout.current)  { clearTimeout(nudgeHoldTimeout.current);   nudgeHoldTimeout.current  = null; }
    if (nudgeHoldInterval.current) { clearInterval(nudgeHoldInterval.current); nudgeHoldInterval.current = null; }
  };

  const doCutAudio = async () => {
    const list  = entriesRef.current;
    const i     = idxRef.current;
    const entry = list[i];
    if (!entry || cutIn === null || cutOut === null || cutIn >= cutOut) {
      setStatus('Set both In and Out points first.'); return;
    }
    const uri = findAudioUri(entry.filename, audioMapRef.current);
    if (!uri) { setStatus('Audio not found.'); return; }

    setLoading(true);
    try {
      const resp   = await fetch(uri);
      const buffer = await resp.arrayBuffer();
      let   newBuf: ArrayBuffer;

      const stem     = entry.filename.replace(/\.[^.]+$/, '');
      const saveName = stem + '_trim.wav';

      if (Platform.OS === 'web') {
        const AudioCtx = (window as any).AudioContext || (window as any).webkitAudioContext;
        if (!AudioCtx) throw new Error('AudioContext not supported');
        const ctx     = new AudioCtx() as AudioContext;
        const decoded = await ctx.decodeAudioData(buffer.slice(0));
        const sr      = decoded.sampleRate;
        const inS     = Math.floor(cutIn  / 1000 * sr);
        const outS    = Math.floor(cutOut / 1000 * sr);
        const newLen  = decoded.length - (outS - inS);
        const nb      = ctx.createBuffer(decoded.numberOfChannels, newLen, sr);
        for (let ch = 0; ch < decoded.numberOfChannels; ch++) {
          const src = decoded.getChannelData(ch);
          const dst = nb.getChannelData(ch);
          dst.set(src.subarray(0, inS), 0);
          dst.set(src.subarray(outS),   inS);
        }
        newBuf = audioBufferToWav(nb);
        const blob = new Blob([newBuf], { type: 'audio/wav' });
        const url  = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url; link.download = saveName; link.click();
        URL.revokeObjectURL(url);
        setStatus(`Cut saved as ${saveName} — downloaded ✓`);
      } else {
        if (!entry.filename.toLowerCase().endsWith('.wav')) {
          throw new Error('Audio cut only supports WAV files on Android.');
        }
        newBuf = cutWavBuffer(buffer, cutIn, cutOut);
        const u8  = new Uint8Array(newBuf);
        let   bin = '';
        const chunk = 8192;
        for (let j = 0; j < u8.length; j += chunk)
          bin += String.fromCharCode(...u8.subarray(j, Math.min(j + chunk, u8.length)));
        const b64 = btoa(bin);
        const folderUri = configRef.current.audioFolderUri;
        if (!folderUri) throw new Error('Audio folder not set — re-pick the audio folder.');
        const trimLower = saveName.toLowerCase();
        const existingUri = audioMapRef.current[trimLower];
        const newFileUri = existingUri
          ? existingUri
          : await LegacyFS.StorageAccessFramework.createFileAsync(folderUri, saveName, 'audio/wav');
        await LegacyFS.StorageAccessFramework.writeAsStringAsync(newFileUri, b64, { encoding: LegacyFS.EncodingType.Base64 });
        audioMapRef.current[trimLower] = newFileUri;
        cutUndoRef.current = { globalIdx: i, oldFilename: entry.filename };
        await persistFilename(i, saveName);
        setStatus(`Saved as ${saveName} ✓  (↩ to undo)`);
      }

      addLog(`✂ ${entry.filename} → ${saveName}: removed ${fmtTime(cutOut - cutIn)} @ ${fmtTime(cutIn)}`);
      setCutIn(null); setCutInText('');
      setCutOut(null); setCutOutText('');
      setCutMode(false);
      await playCurrentEntry();
    } catch (e) {
      setStatus(`Cut error: ${String(e)}`);
    } finally {
      setLoading(false);
    }
  };

  const applyReplaceAll = async () => {
    const wb    = workbookRef.current;
    const sheet = curSheetRef.current;
    const all   = allEntriesRef.current;
    if (!wb || !sheet || !all.length || !filterInput.trim() || !replaceInput.trim()) return;
    const fi = filteredIndicesRef.current;
    if (!fi.length) return;

    const patterns = filterInput.split(',').map(p => p.trim()).filter(Boolean);
    const ws       = wb.Sheets[sheet];
    const rows     = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 });

    const dataRowMap: number[] = [];
    let dataIdx = 0;
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r] as unknown[];
      if (Array.isArray(row) && row.length >= 2 && row[0] && row[1]) {
        dataRowMap[dataIdx++] = r;
      }
    }

    const snapshot: { globalIdx: number; oldText: string }[] = [];
    let count = 0;
    for (const globalIdx of fi) {
      const entry = all[globalIdx];
      if (!entry) continue;
      let newText = entry.text;
      for (const pat of patterns) newText = applyReplacement(newText, pat, replaceInput.trim());
      if (newText !== entry.text) {
        snapshot.push({ globalIdx, oldText: entry.text });
        const rowIdx = dataRowMap[globalIdx];
        if (rowIdx !== undefined)
          ws[XLSX.utils.encode_cell({ r: rowIdx, c: 0 })] = { t: 's', v: newText.trim(), w: newText.trim() };
        all[globalIdx]                = { ...entry, text: newText.trim() };
        entriesRef.current[globalIdx] = { ...entry, text: newText.trim() };
        count++;
      }
    }
    undoRef.current = snapshot.length > 0 ? snapshot : null;

    setEntries([...entriesRef.current]);
    const cur = entriesRef.current[idxRef.current];
    if (cur) setEditText(cur.text);

    const lp = xlsxLocalRef.current;
    const msg = `Replace "${filterInput}" → "${replaceInput}" in ${count} entr${count === 1 ? 'y' : 'ies'}`;
    if (lp && Platform.OS !== 'web') {
      const base64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx', compression: true });
      try {
        await LegacyFS.writeAsStringAsync(lp, base64, { encoding: LegacyFS.EncodingType.Base64 });
        setStatus(`Replaced in ${count} entr${count === 1 ? 'y' : 'ies'} ✓`);
        addLog(msg);
      } catch (e) { setStatus(`Replace error: ${String(e)}`); }
    } else {
      setStatus(`Replaced in ${count} entr${count === 1 ? 'y' : 'ies'} ✓`);
      addLog(msg);
    }
  };

  const applyUndo = async () => {
    // Cut undo: revert filename back to original
    const cutSnap = cutUndoRef.current;
    if (cutSnap) {
      cutUndoRef.current = null;
      delete audioMapRef.current[entriesRef.current[cutSnap.globalIdx]?.filename.toLowerCase() ?? ''];
      await persistFilename(cutSnap.globalIdx, cutSnap.oldFilename);
      setStatus(`Cut undone — reverted to ${cutSnap.oldFilename} ✓`);
      addLog(`Undo cut: restored ${cutSnap.oldFilename}`);
      await playCurrentEntry();
      return;
    }
    const snap = undoRef.current;
    const wb   = workbookRef.current;
    const sheet = curSheetRef.current;
    if (!snap || !snap.length || !wb || !sheet) { setStatus('Nothing to undo.'); return; }
    const all = allEntriesRef.current;
    const ws  = wb.Sheets[sheet];
    const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 });
    const dataRowMap: number[] = [];
    let di = 0;
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r] as unknown[];
      if (Array.isArray(row) && row.length >= 2 && row[0] && row[1]) dataRowMap[di++] = r;
    }
    for (const { globalIdx, oldText } of snap) {
      const rowIdx = dataRowMap[globalIdx];
      if (rowIdx !== undefined)
        ws[XLSX.utils.encode_cell({ r: rowIdx, c: 0 })] = { t: 's', v: oldText, w: oldText };
      all[globalIdx]                = { ...all[globalIdx], text: oldText };
      entriesRef.current[globalIdx] = { ...entriesRef.current[globalIdx], text: oldText };
    }
    undoRef.current = null;
    setEntries([...entriesRef.current]);
    const cur = entriesRef.current[idxRef.current];
    if (cur) setEditText(cur.text);
    const lp = xlsxLocalRef.current;
    if (lp && Platform.OS !== 'web') {
      const base64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx', compression: true });
      try {
        await LegacyFS.writeAsStringAsync(lp, base64, { encoding: LegacyFS.EncodingType.Base64 });
        setStatus(`Undone (${snap.length} entr${snap.length === 1 ? 'y' : 'ies'}) ✓`);
        addLog(`Undo: restored ${snap.length} entr${snap.length === 1 ? 'y' : 'ies'}`);
      } catch (e) { setStatus(`Undo save error: ${String(e)}`); }
    } else {
      setStatus(`Undone (${snap.length} entr${snap.length === 1 ? 'y' : 'ies'}) ✓`);
      addLog(`Undo: restored ${snap.length} entr${snap.length === 1 ? 'y' : 'ies'}`);
    }
  };

  const hasEntries = entries.length > 0;

  return (
    <KeyboardAvoidingView
      style={s.root}
      behavior="padding"
    >
      <StatusBar style="light" />

      <View style={s.row}>
        <Btn label="Excel"  onPress={pickExcel}       bg="#4fc3f7" fg={C.bg} />
        <Btn label="Audio"  onPress={pickAudioFolder} bg="#81d4fa" fg={C.bg} />
        <Btn label="Share"  onPress={doShare}         bg="#01579b" fg="#fff" />
        <TouchableOpacity style={s.aboutBtn} onPress={() => setAboutModal(true)}>
          <Text style={s.aboutBtnText}>?</Text>
        </TouchableOpacity>
      </View>

      <Text style={s.status} numberOfLines={2}>
        {xlsxLabel ? `${xlsxLabel}  ` : ''}{status}{audioCount > 0 ? `  [${audioCount} audio]` : ''}
      </Text>

      {hasEntries && (
        <View style={s.infoRow}>
          <Text style={s.counter}>{displayIdx + 1}/{filteredCount}</Text>
          <TouchableOpacity style={s.filenameBtn} onPress={() => {
            setEditFilename(entries[idx]?.filename ?? '');
            setFilenameModal(true);
          }}>
            <Text style={s.filename} numberOfLines={1}>{entries[idx]?.filename}</Text>
          </TouchableOpacity>
          {duration > 0 && <Text style={s.sheetTag}>{fmtTime(duration)}</Text>}
          {curSheet && <Text style={s.sheetTag} numberOfLines={1}>{curSheet}</Text>}
        </View>
      )}

      <TextInput
        style={s.textInput}
        value={editText}
        onChangeText={setEditText}
        multiline
        placeholder="Load audio folder and Excel file to begin."
        placeholderTextColor={C.muted}
      />

      {/* Speed — hold to repeat */}
      <View style={s.sliderRow}>
        <Text style={s.muted}>Speed:</Text>
        <TouchableOpacity style={s.speedBtn} onPressIn={() => startSpeedHold(-0.05)} onPressOut={stopSpeedHold}>
          <Text style={s.speedBtnText}>−</Text>
        </TouchableOpacity>
        <Text style={s.accentLabel}>{speed.toFixed(2)}×</Text>
        <TouchableOpacity style={s.speedBtn} onPressIn={() => startSpeedHold(0.05)} onPressOut={stopSpeedHold}>
          <Text style={s.speedBtnText}>+</Text>
        </TouchableOpacity>
        <TouchableOpacity style={s.undoBtn} onPress={applyUndo}>
          <Text style={s.speedBtnText}>↩</Text>
        </TouchableOpacity>
        {hasEntries && (
          <TouchableOpacity
            style={[s.undoBtn, cutMode && { backgroundColor: C.red }]}
            onPress={() => {
              const next = !cutMode;
              setCutMode(next);
              if (next) {
                speedRef.current = 0.75; setSpeed(0.75);
                soundRef.current?.setRateAsync(0.75, false).catch(() => {});
              }
            }}
          >
            <Text style={s.speedBtnText}>✂</Text>
          </TouchableOpacity>
        )}
      </View>

      {/* Play button + progress bar */}
      <View style={s.progressRow}>
        <TouchableOpacity style={s.playBtn} onPress={() => playCurrentEntry()}>
          <Text style={s.playBtnText}>{playing ? '▶▶' : '▶'}</Text>
        </TouchableOpacity>
        <View
          style={s.progressTrack}
          onLayout={(e) => { trackWidthRef.current = e.nativeEvent.layout.width || 1; }}
          onStartShouldSetResponder={() => true}
          onResponderGrant={(e) => {
            (async () => {
              if (!soundRef.current || duration === 0) {
                await playCurrentEntry();
                return;
              }
              const ratio  = Math.max(0, Math.min(1, e.nativeEvent.locationX / trackWidthRef.current));
              const seekMs = Math.floor(ratio * duration);
              await soundRef.current.setPositionAsync(seekMs);
              const st = await soundRef.current.getStatusAsync();
              if (st.isLoaded && !st.isPlaying) {
                await soundRef.current.playAsync();
                setPlaying(true);
              }
            })().catch(() => {});
          }}
        >
          {!playingPreview && <View style={[s.progressFill, { width: `${duration > 0 ? Math.min(position / duration * 100, 100) : 0}%` }]} />}
          {!!editText && (() => {
            const words = editText.trim().split(/\s+/).filter(Boolean);
            const totalChars = words.join('').length;
            const fontSize = Math.max(6, Math.min(10, Math.floor((trackWidthRef.current * 0.85) / Math.max(1, totalChars * 0.6))));
            return (
              <View style={s.progressWords} pointerEvents="none">
                {words.map((word, i) => (
                  <Text key={i} style={[s.progressWordText, { fontSize }]}>{word}</Text>
                ))}
              </View>
            );
          })()}
        </View>
      </View>

      {cutMode && hasEntries && (() => {
        const canCut = cutIn !== null && cutOut !== null && cutIn < cutOut!;
        return (
          <View style={s.progressRow}>
            <TouchableOpacity style={[s.playBtn, { backgroundColor: '#01579b' }]} onPress={playBottomBand}>
              <Text style={s.playBtnText}>{canCut ? '✂▶' : '▶'}</Text>
            </TouchableOpacity>
            <View style={[s.progressTrack, { backgroundColor: '#0a1929' }]}>
              {playingPreview && (
                <View style={[s.progressFill, { width: `${duration > 0 ? Math.min(position / duration * 100, 100) : 0}%`, backgroundColor: '#01579b' }]} />
              )}
              {canCut && !playingPreview && duration > 0 && (
                <View pointerEvents="none" style={{
                  position: 'absolute', top: 0, bottom: 0,
                  left:  `${Math.min(cutIn!  / duration * 100, 100)}%` as any,
                  width: `${Math.min((cutOut! - cutIn!) / duration * 100, 100)}%` as any,
                  backgroundColor: 'rgba(239,83,80,0.55)',
                }} />
              )}
              <Text pointerEvents="none" style={{ position: 'absolute', bottom: 4, left: 8, color: canCut ? '#fff' : C.muted, fontSize: 11 }}>
                {canCut
                  ? `${fmtTime(cutIn!)} → ${fmtTime(cutOut!)}   −${fmtTime(cutOut! - cutIn!)}`
                  : 'press ▶ to play, then mark In & Out'}
              </Text>
            </View>
          </View>
        );
      })()}

      {cutMode && hasEntries && (() => {
        const canAct = cutIn !== null && cutOut !== null && cutIn < cutOut;
        return (
          <>
            <View style={s.cutRow}>
              <TouchableOpacity style={[s.cutMarkBtn, cutIn !== null && { backgroundColor: C.red }]} onPress={markIn}>
                <Text style={s.cutMarkText}>● In</Text>
              </TouchableOpacity>
              <TouchableOpacity style={s.nudgeBtn} onPressIn={() => startNudgeHold('in', -10)} onPressOut={stopNudgeHold}>
                <Text style={s.nudgeBtnText}>◀</Text>
              </TouchableOpacity>
              <TextInput style={s.cutTimeInput} value={cutInText} onChangeText={onCutInTextChange}
                keyboardType="decimal-pad" placeholder="sec" placeholderTextColor={C.muted} />
              <TouchableOpacity style={s.nudgeBtn} onPressIn={() => startNudgeHold('in', 10)} onPressOut={stopNudgeHold}>
                <Text style={s.nudgeBtnText}>▶</Text>
              </TouchableOpacity>
              <TouchableOpacity style={[s.cutMarkBtn, cutOut !== null && { backgroundColor: C.red }]} onPress={markOut}>
                <Text style={s.cutMarkText}>● Out</Text>
              </TouchableOpacity>
              <TouchableOpacity style={s.nudgeBtn} onPressIn={() => startNudgeHold('out', -10)} onPressOut={stopNudgeHold}>
                <Text style={s.nudgeBtnText}>◀</Text>
              </TouchableOpacity>
              <TextInput style={s.cutTimeInput} value={cutOutText} onChangeText={onCutOutTextChange}
                keyboardType="decimal-pad" placeholder="sec" placeholderTextColor={C.muted} />
              <TouchableOpacity style={s.nudgeBtn} onPressIn={() => startNudgeHold('out', 10)} onPressOut={stopNudgeHold}>
                <Text style={s.nudgeBtnText}>▶</Text>
              </TouchableOpacity>
              <TouchableOpacity style={[s.cutApplyBtn, !canAct && { opacity: 0.35 }]} onPress={doCutAudio}>
                <Text style={s.cutMarkText}>✂  Cut</Text>
              </TouchableOpacity>
            </View>
          </>
        );
      })()}

      {/* Swipe strip with arrow buttons on edges */}
      <View style={s.swipeStrip} {...swipeResponder.panHandlers}>
        <TouchableOpacity style={s.swipeArrowBtn} onPress={goPrev}>
          <Text style={s.swipeArrowText}>◀</Text>
        </TouchableOpacity>
        <Text style={s.swipeHint}>swipe to navigate</Text>
        <TouchableOpacity style={s.swipeArrowBtn} onPress={goNext}>
          <Text style={s.swipeArrowText}>▶</Text>
        </TouchableOpacity>
      </View>

      {/* Find / Go / Replace on one row */}
      <View style={s.filterJumpRow}>
        <Text style={s.muted}>Find:</Text>
        <TextInput
          style={s.findInput}
          value={filterInput}
          onChangeText={setFilterInput}
          returnKeyType="search"
          onSubmitEditing={applyFilter}
          placeholder="e.g. p-, -a-, -ng, or r"
          placeholderTextColor={C.muted}
        />
        <Text style={s.muted}>Go:</Text>
        <TextInput
          style={s.jumpInput}
          value={jumpText}
          onChangeText={setJumpText}
          keyboardType="number-pad"
          returnKeyType="done"
          onEndEditing={doJump}
          placeholder="#"
          placeholderTextColor={C.muted}
        />
        <Text style={s.muted}>Replace:</Text>
        <TextInput
          style={s.smallInput}
          value={replaceInput}
          onChangeText={setReplaceInput}
          returnKeyType="done"
          onSubmitEditing={applyReplaceAll}
          placeholderTextColor={C.muted}
        />
      </View>

      <Modal visible={sheetModal} transparent animationType="fade"
             onRequestClose={() => setSheetModal(false)}>
        <View style={s.overlay}>
          <View style={s.modalBox}>
            <Text style={s.modalTitle}>Choose Sheet</Text>
            <ScrollView>
              {sheets.map((sh) => (
                <TouchableOpacity key={sh} style={s.sheetItem} onPress={() => loadSheet(sh)}>
                  <Text style={s.sheetItemText}>{sh}</Text>
                </TouchableOpacity>
              ))}
            </ScrollView>
            <Btn label="Cancel" onPress={() => setSheetModal(false)} bg={C.red} />
          </View>
        </View>
      </Modal>

      <Modal visible={filenameModal} transparent animationType="fade"
             onRequestClose={() => setFilenameModal(false)}>
        <KeyboardAvoidingView style={{ flex: 1 }} behavior="padding">
          <View style={s.overlay}>
            <View style={s.modalBox}>
              <Text style={s.modalTitle}>Edit Filename</Text>
              <TextInput
                style={s.modalInput}
                value={editFilename}
                onChangeText={setEditFilename}
                autoFocus
                selectTextOnFocus
                returnKeyType="done"
                onSubmitEditing={() => persistFilename(idxRef.current, editFilename)}
              />
              <View style={s.row}>
                <Btn label="Cancel"  onPress={() => setFilenameModal(false)} bg={C.red} />
                <Btn label="Confirm" onPress={() => persistFilename(idxRef.current, editFilename)} bg={C.green} fg="#fff" />
              </View>
              <TouchableOpacity style={s.deleteEntryBtn} onPress={() => deleteEntry(idxRef.current)}>
                <Text style={s.deleteEntryText}>🗑 Delete this entry</Text>
              </TouchableOpacity>
            </View>
          </View>
        </KeyboardAvoidingView>
      </Modal>

      <Modal visible={aboutModal} transparent animationType="fade"
             onRequestClose={() => setAboutModal(false)}>
        <View style={s.overlay}>
          <View style={s.modalBox}>
            <Text style={s.modalTitle}>About</Text>
            <ScrollView>
              <Text style={s.aboutText}>
                An app for reviewing audio transcriptions. Pick an Excel file first, then an audio folder — the folder name auto-selects the matching sheet (e.g. a "words" folder loads the words sheet, "sentences" loads the sentences sheet).{'\n\n'}
                Click the progress bar to re-listen to any part of the audio.{'\n\n'}
                Find filters by sound position: p (anywhere), p- (word-start), -p (word-end), -p- (middle). Comma-separate for OR logic — e.g. p-,f- finds all entries with a word starting with p or f.{'\n\n'}
                Replace applies a substitution to all filtered entries. ↩ undoes the last bulk replace.{'\n\n'}
                Feedback: shidailun@gmail.com
              </Text>
            </ScrollView>
            <Btn label="Close" onPress={() => setAboutModal(false)} bg={C.accent} fg={C.bg} />
          </View>
        </View>
      </Modal>

      {logLines.length > 0 && (
        <View style={s.logContainer}>
          <View style={s.logHeader}>
            <Text style={s.logHeaderText}>Log</Text>
            <TouchableOpacity onPress={async () => {
              const content = logLines.join('\n');
              if (Platform.OS === 'web') {
                const blob = new Blob([content], { type: 'text/plain' });
                const url  = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url; link.download = 'transcription_log.txt'; link.click();
                URL.revokeObjectURL(url);
              } else {
                const path = (LegacyFS.documentDirectory ?? '') + 'transcription_log.txt';
                await LegacyFS.writeAsStringAsync(path, content);
                await Sharing.shareAsync(path, { mimeType: 'text/plain', dialogTitle: 'Share log' });
              }
            }}>
              <Text style={s.logShareBtn}>↑ Share</Text>
            </TouchableOpacity>
          </View>
          <ScrollView style={s.logScroll} contentContainerStyle={{ paddingVertical: 4 }}>
            {logLines.map((line, i) => (
              <Text key={i} style={s.logLine}>{line}</Text>
            ))}
          </ScrollView>
        </View>
      )}

      {loading && (
        <View style={s.loadingOverlay}>
          <ActivityIndicator size="large" color={C.accent} />
          <Text style={{ color: C.text, marginTop: 10 }}>Loading…</Text>
        </View>
      )}
    </KeyboardAvoidingView>
  );
}

function crc32(data: Uint8Array): number {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) {
    c ^= data[i];
    for (let j = 0; j < 8; j++) c = (c >>> 1) ^ (c & 1 ? 0xEDB88320 : 0);
  }
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function buildZip(files: { name: string; data: Uint8Array }[]): Uint8Array {
  const enc = new TextEncoder();
  const locals: Uint8Array[] = [];
  const centrals: Uint8Array[] = [];
  let off = 0;
  for (const f of files) {
    const nb = enc.encode(f.name);
    const crc = crc32(f.data);
    const sz  = f.data.length;
    const lh  = new Uint8Array(30 + nb.length);
    const lv  = new DataView(lh.buffer);
    lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true);
    lv.setUint32(14, crc, true); lv.setUint32(18, sz, true); lv.setUint32(22, sz, true);
    lv.setUint16(26, nb.length, true); lh.set(nb, 30);
    const ch = new Uint8Array(46 + nb.length);
    const cv = new DataView(ch.buffer);
    cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true);
    cv.setUint32(16, crc, true); cv.setUint32(20, sz, true); cv.setUint32(24, sz, true);
    cv.setUint16(28, nb.length, true); cv.setUint32(42, off, true); ch.set(nb, 46);
    locals.push(lh, f.data); centrals.push(ch);
    off += lh.length + sz;
  }
  const cdSize = centrals.reduce((s, c) => s + c.length, 0);
  const eocd = new Uint8Array(22);
  const ev   = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, files.length, true);
  ev.setUint16(10, files.length, true); ev.setUint32(12, cdSize, true); ev.setUint32(16, off, true);
  const parts = [...locals, ...centrals, eocd];
  const total = parts.reduce((s, p) => s + p.length, 0);
  const out = new Uint8Array(total); let pos = 0;
  for (const p of parts) { out.set(p, pos); pos += p.length; }
  return out;
}

function insertXlsxRowAfter(ws: XLSX.WorkSheet, afterRow: number, col0: string, col1: string) {
  const range = XLSX.utils.decode_range(ws['!ref'] ?? 'A1');
  for (let r = range.e.r; r > afterRow; r--)
    for (let c = 0; c <= range.e.c; c++) {
      const from = XLSX.utils.encode_cell({ r, c });
      const to   = XLSX.utils.encode_cell({ r: r + 1, c });
      if (ws[from]) { ws[to] = ws[from]; delete ws[from]; } else { delete ws[to]; }
    }
  const nr = afterRow + 1;
  ws[XLSX.utils.encode_cell({ r: nr, c: 0 })] = { t: 's', v: col0, w: col0 };
  ws[XLSX.utils.encode_cell({ r: nr, c: 1 })] = { t: 's', v: col1, w: col1 };
  ws['!ref'] = XLSX.utils.encode_range({ s: range.s, e: { r: range.e.r + 1, c: range.e.c } });
}

function removeXlsxRow(ws: XLSX.WorkSheet, rowIdx: number) {
  const range = XLSX.utils.decode_range(ws['!ref'] ?? 'A1');
  for (let c = 0; c <= range.e.c; c++) delete ws[XLSX.utils.encode_cell({ r: rowIdx, c })];
  for (let r = rowIdx + 1; r <= range.e.r; r++)
    for (let c = 0; c <= range.e.c; c++) {
      const from = XLSX.utils.encode_cell({ r, c });
      const to   = XLSX.utils.encode_cell({ r: r - 1, c });
      if (ws[from]) { ws[to] = ws[from]; delete ws[from]; } else { delete ws[to]; }
    }
  if (range.e.r > range.s.r)
    ws['!ref'] = XLSX.utils.encode_range({ s: range.s, e: { r: range.e.r - 1, c: range.e.c } });
}

function cutWavBuffer(buffer: ArrayBuffer, inMs: number, outMs: number): ArrayBuffer {
  const view = new DataView(buffer);
  const tag  = (o: number) =>
    String.fromCharCode(view.getUint8(o), view.getUint8(o+1), view.getUint8(o+2), view.getUint8(o+3));
  if (tag(0) !== 'RIFF' || tag(8) !== 'WAVE') throw new Error('Not a WAV file');

  let off = 12, sampleRate = 0, numCh = 0, bps = 0, dataOff = -1, dataSize = 0;
  while (off + 8 <= buffer.byteLength) {
    const id   = tag(off);
    const size = view.getUint32(off + 4, true);
    if (id === 'fmt ') {
      numCh      = view.getUint16(off + 10, true);
      sampleRate = view.getUint32(off + 12, true);
      bps        = view.getUint16(off + 22, true);
    } else if (id === 'data') {
      dataOff  = off + 8;
      dataSize = size;
      break;
    }
    off += 8 + size + (size % 2);
  }
  if (dataOff < 0 || sampleRate === 0) throw new Error('Could not parse WAV');

  const blockAlign = numCh * Math.ceil(bps / 8);
  const inByte     = Math.min(Math.floor(inMs  / 1000 * sampleRate) * blockAlign, dataSize);
  const outByte    = Math.min(Math.floor(outMs / 1000 * sampleRate) * blockAlign, dataSize);
  if (inByte >= outByte) throw new Error('Cut range is empty');

  const newDataSize = dataSize - (outByte - inByte);
  const newBuf      = new ArrayBuffer(dataOff + newDataSize);
  const nb          = new Uint8Array(newBuf);
  const nv          = new DataView(newBuf);

  nb.set(new Uint8Array(buffer, 0, dataOff));
  nv.setUint32(4,          newBuf.byteLength - 8, true);
  nv.setUint32(dataOff - 4, newDataSize,           true);
  nb.set(new Uint8Array(buffer, dataOff, inByte),                        dataOff);
  nb.set(new Uint8Array(buffer, dataOff + outByte, dataSize - outByte),  dataOff + inByte);

  // 5ms crossfade at splice point to eliminate click (16-bit PCM only)
  if (bps === 16) {
    const fadeSamples = Math.min(
      Math.floor(0.005 * sampleRate),
      Math.floor(Math.min(inByte, dataSize - outByte) / blockAlign / 2),
    );
    const nv2 = new DataView(newBuf);
    for (let s = 0; s < fadeSamples; s++) {
      const fadeOut = (fadeSamples - 1 - s) / fadeSamples;
      const fadeIn  = s / fadeSamples;
      for (let ch = 0; ch < numCh; ch++) {
        const oOut = dataOff + inByte - (fadeSamples - s) * blockAlign + ch * 2;
        const oIn  = dataOff + inByte + s * blockAlign + ch * 2;
        if (oOut >= dataOff) nv2.setInt16(oOut, Math.round(nv2.getInt16(oOut, true) * fadeOut), true);
        if (oIn  < dataOff + newDataSize) nv2.setInt16(oIn,  Math.round(nv2.getInt16(oIn,  true) * fadeIn),  true);
      }
    }
  }

  return newBuf;
}

function audioBufferToWav(buf: AudioBuffer): ArrayBuffer {
  const numCh  = buf.numberOfChannels;
  const sr     = buf.sampleRate;
  const len    = buf.length;
  const block  = numCh * 2;
  const data   = new ArrayBuffer(44 + len * block);
  const view   = new DataView(data);
  const ws     = (o: number, s: string) => { for (let i = 0; i < s.length; i++) view.setUint8(o + i, s.charCodeAt(i)); };
  ws(0, 'RIFF'); view.setUint32(4, data.byteLength - 8, true); ws(8, 'WAVE');
  ws(12, 'fmt '); view.setUint32(16, 16, true); view.setUint16(20, 1, true);
  view.setUint16(22, numCh, true); view.setUint32(24, sr, true);
  view.setUint32(28, sr * block, true); view.setUint16(32, block, true); view.setUint16(34, 16, true);
  ws(36, 'data'); view.setUint32(40, len * block, true);
  const out = new Int16Array(data, 44);
  for (let s = 0; s < len; s++)
    for (let ch = 0; ch < numCh; ch++) {
      const v = Math.max(-1, Math.min(1, buf.getChannelData(ch)[s]));
      out[s * numCh + ch] = v < 0 ? v * 32768 : v * 32767;
    }
  return data;
}

function wordDiff(oldText: string, newText: string): string {
  if (oldText === newText) return '';
  const a = oldText.trim().split(/\s+/);
  const b = newText.trim().split(/\s+/);
  let start = 0;
  while (start < a.length && start < b.length && a[start] === b[start]) start++;
  let endA = a.length - 1, endB = b.length - 1;
  while (endA > start && endB > start && a[endA] === b[endB]) { endA--; endB--; }
  const changedA = a.slice(start, endA + 1).join(' ');
  const changedB = b.slice(start, endB + 1).join(' ');
  if (!changedA && !changedB) return '';
  return `${changedA || '∅'}→${changedB || '∅'}`;
}

function logTime(): string {
  const d = new Date();
  return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
}

function fmtTime(ms: number): string {
  const tenths = Math.floor(ms / 100);
  const s      = Math.floor(tenths / 10);
  return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}.${tenths % 10}`;
}

function Btn({
  label, onPress, bg = C.surface, fg = C.text, flex = 1,
}: { label: string; onPress: () => void; bg?: string; fg?: string; flex?: number }) {
  return (
    <TouchableOpacity style={[s.btn, { backgroundColor: bg, flex }]}
                      onPress={onPress} activeOpacity={0.75}>
      <Text style={[s.btnText, { color: fg }]}>{label}</Text>
    </TouchableOpacity>
  );
}

const s = StyleSheet.create({
  root:          { flex: 1, backgroundColor: C.bg, padding: 8, paddingTop: 48, paddingBottom: 16 },
  row:           { flexDirection: 'row', gap: 4, marginBottom: 6 },
  btn:           { height: 52, borderRadius: 8, justifyContent: 'center', alignItems: 'center', paddingHorizontal: 4 },
  btnText:       { fontSize: 14, fontWeight: '700' },
  status:        { color: C.muted, fontSize: 12, marginBottom: 4, minHeight: 34 },
  infoRow:       { flexDirection: 'row', alignItems: 'center', marginBottom: 4, gap: 8 },
  counter:       { color: C.accent, fontSize: 14, fontWeight: '700' },
  filenameBtn:   { flex: 1 },
  filename:      { color: C.muted, fontSize: 11 },
  sheetTag:      { color: C.muted, fontSize: 10, maxWidth: 100 },
  textInput:     {
    height: 80, backgroundColor: C.surface, color: C.text,
    fontSize: 17, padding: 8, borderRadius: 8, marginBottom: 4, textAlignVertical: 'top',
    borderWidth: 1, borderColor: C.muted,
  },
  progressRow:   { flexDirection: 'row', alignItems: 'center', gap: 6, marginBottom: 4 },
  playBtn:       { width: 52, height: 60, backgroundColor: C.accent, borderRadius: 8, justifyContent: 'center', alignItems: 'center' },
  playBtnText:   { color: C.bg, fontSize: 16, fontWeight: '700' },
  progressTrack: { flex: 1, height: 60, backgroundColor: C.surface, borderRadius: 6, overflow: 'hidden' },
  progressFill:  { height: '100%', backgroundColor: C.accent, borderRadius: 6 },
  progressWords: { position: 'absolute', left: 6, right: 6, top: 0, bottom: 0, flexDirection: 'row', justifyContent: 'space-between', alignItems: 'center' },
  progressWordText: { color: '#fff', textShadowColor: 'rgba(0,0,0,0.9)', textShadowOffset: { width: 0, height: 0 }, textShadowRadius: 4 },
  sliderRow:     { flexDirection: 'row', alignItems: 'center', height: 44, marginBottom: 4, gap: 8 },
  speedBtn:      { width: 44, height: 44, backgroundColor: C.surface, borderRadius: 8, justifyContent: 'center', alignItems: 'center' },
  undoBtn:       { width: 44, height: 44, backgroundColor: C.red, borderRadius: 8, justifyContent: 'center', alignItems: 'center' },
  speedBtnText:  { color: C.text, fontSize: 22, fontWeight: '700' },
  muted:         { color: C.muted, fontSize: 13 },
  accentLabel:   { color: C.accent, fontSize: 13, textAlign: 'center', paddingHorizontal: 2 },
  swipeStrip:    { height: 60, flexDirection: 'row', justifyContent: 'space-between', alignItems: 'center', backgroundColor: C.surface, borderRadius: 8, marginBottom: 4, paddingHorizontal: 12 },
  swipeHint:     { color: C.muted, fontSize: 12 },
  swipeArrowBtn: { padding: 8 },
  swipeArrowText:{ color: C.muted, fontSize: 18 },
  filterJumpRow: { flexDirection: 'row', alignItems: 'center', gap: 6 },
  findInput:     { flex: 1, height: 44, backgroundColor: C.surface, color: C.text, borderRadius: 8, paddingHorizontal: 10, fontSize: 14 },
  smallInput:    { width: 60, height: 44, backgroundColor: C.surface, color: C.text, borderRadius: 8, paddingHorizontal: 8, fontSize: 14 },
  jumpInput:     { width: 60, height: 44, backgroundColor: C.surface, color: C.text, borderRadius: 8, paddingHorizontal: 10, fontSize: 14, textAlign: 'center' },
  modalInput:    { backgroundColor: C.bg, color: C.text, borderRadius: 8, paddingHorizontal: 12, paddingVertical: 10, fontSize: 15, marginBottom: 12, borderWidth: 1, borderColor: C.accent },
  overlay:       { flex: 1, backgroundColor: 'rgba(0,0,0,0.65)', justifyContent: 'center', alignItems: 'center' },
  modalBox:      { backgroundColor: C.surface, borderRadius: 12, padding: 16, width: '82%', maxHeight: '70%' },
  modalTitle:    { color: C.text, fontSize: 16, fontWeight: '700', marginBottom: 12 },
  sheetItem:     { padding: 14, borderRadius: 8, backgroundColor: C.bg, marginBottom: 6 },
  sheetItemText: { color: C.text, fontSize: 14 },
  loadingOverlay:{ ...StyleSheet.absoluteFillObject, backgroundColor: 'rgba(0,0,0,0.55)', justifyContent: 'center', alignItems: 'center' },
  logContainer:  { marginTop: 6, backgroundColor: C.surface, borderRadius: 8 },
  logHeader:     { flexDirection: 'row', justifyContent: 'space-between', alignItems: 'center', paddingHorizontal: 10, paddingTop: 6, paddingBottom: 2 },
  logHeaderText: { color: C.muted, fontSize: 11, fontWeight: '700', textTransform: 'uppercase', letterSpacing: 1 },
  logShareBtn:   { color: '#01579b', fontSize: 12, fontWeight: '700' },
  logScroll:     { maxHeight: 99, paddingHorizontal: 10 },
  logLine:       { color: C.muted, fontSize: 11, fontFamily: Platform.OS === 'ios' ? 'Menlo' : 'monospace', paddingVertical: 2 },
  deleteEntryBtn:{ marginTop: 12, alignItems: 'center', padding: 10 },
  deleteEntryText:{ color: C.red, fontSize: 13, fontWeight: '700' },
  aboutBtn:      { width: 44, height: 52, backgroundColor: C.surface, borderRadius: 8, justifyContent: 'center', alignItems: 'center' },
  aboutBtnText:  { color: C.muted, fontSize: 18, fontWeight: '700' },
  aboutText:     { color: C.text, fontSize: 14, lineHeight: 22 },
cutRow:        { flexDirection: 'row', alignItems: 'center', gap: 3, marginBottom: 4 },
  nudgeBtn:      { width: 28, height: 44, backgroundColor: C.surface, borderRadius: 6, justifyContent: 'center', alignItems: 'center' },
  nudgeBtnText:  { color: C.muted, fontSize: 12 },
  cutMarkBtn:    { height: 44, paddingHorizontal: 8, backgroundColor: '#b71c1c', borderRadius: 8, justifyContent: 'center', alignItems: 'center' },
  cutTimeInput:  { width: 54, height: 44, backgroundColor: C.bg, color: C.text, borderRadius: 8, paddingHorizontal: 6, fontSize: 13, textAlign: 'center', borderWidth: 1, borderColor: C.red },
  cutPreviewBtn: { flex: 1, height: 44, paddingHorizontal: 6, backgroundColor: '#01579b', borderRadius: 8, justifyContent: 'center', alignItems: 'center' },
  cutApplyBtn:   { flex: 2, height: 44, paddingHorizontal: 6, backgroundColor: C.red,    borderRadius: 8, justifyContent: 'center', alignItems: 'center' },
  cutMarkText:   { color: '#fff', fontSize: 11, fontWeight: '700', textAlign: 'center' },
  cutRange:      { flex: 1, color: C.text, fontSize: 12, textAlign: 'center' },
});
