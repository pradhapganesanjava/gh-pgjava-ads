// ── AnkiImporter ──────────────────────────────────────────────────────────────
// Parses .colpkg / .apkg (ZIP archives containing an SQLite DB) in the browser.
// Requires JSZip (loaded via CDN in index.html) + sql.js (loaded via CDN).
// Writes cards to Google Sheets; uploads media to Google Drive via Media module.
//
// Usage:
//   await AnkiImporter.run({ file, deckFilter, mediaMode, token, onProgress })
//   Returns { cards: N, mediaUploaded: M }

window.AnkiImporter = (() => {

  // ── sql.js init ───────────────────────────────────────────────────────────────
  let _SQL = null;
  async function ensureSQL() {
    if (_SQL) return _SQL;
    _SQL = await initSqlJs({
      locateFile: f => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/${f}`
    });
    return _SQL;
  }

  // ── ZIP extraction ────────────────────────────────────────────────────────────
  async function extractZip(file) {
    const zip = await JSZip.loadAsync(file);
    return zip;
  }

  // ── Anki note model helpers ───────────────────────────────────────────────────
  function renderNote(flds, model) {
    const fields = flds.split('\x1f');
    const tmpl = (model.tmpls || [])[0] || {};
    const qfmt = tmpl.qfmt || '';
    const afmt = tmpl.afmt || '';

    function substitute(fmt, fldArr) {
      const fieldNames = (model.flds || []).map(f => f.name);
      let out = fmt;
      fieldNames.forEach((name, i) => {
        out = out.replace(new RegExp(`\\{\\{${name}\\}\\}`, 'g'), fldArr[i] || '');
        out = out.replace(new RegExp(`\\{\\{text:${name}\\}\\}`, 'g'), stripHTML(fldArr[i] || ''));
      });
      out = out.replace(/\{\{FrontSide\}\}/g, substitute(qfmt, fldArr));
      out = out.replace(/\{\{[^}]+\}\}/g, '');
      return out.trim();
    }

    return {
      front: substitute(qfmt, fields) || fields[0] || '',
      back:  substitute(afmt, fields) || fields[1] || ''
    };
  }

  function stripHTML(html) {
    return html.replace(/<[^>]+>/g, '');
  }

  // ── Media filename → MIME type ─────────────────────────────────────────────
  function mimeFor(filename) {
    const ext = (filename.split('.').pop() || '').toLowerCase();
    const map = {
      jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png',
      gif: 'image/gif',  webp: 'image/webp', svg: 'image/svg+xml',
      mp3: 'audio/mpeg', ogg: 'audio/ogg',   wav: 'audio/wav',
      mp4: 'video/mp4',  webm: 'video/webm',
      html: 'text/html', css: 'text/css', js: 'application/javascript'
    };
    return map[ext] || 'application/octet-stream';
  }

  // ── Upload one media file to Google Drive ─────────────────────────────────
  async function uploadToDrive(token, filename, uint8Data, mime, folderId) {
    // Check if already uploaded (cached in Media map)
    if (Media._map?.[filename]?.drive_url) return Media._map[filename];

    const metadata = {
      name: filename,
      mimeType: mime,
      ...(folderId ? { parents: [folderId] } : {})
    };

    const boundary = 'ankiads_bnd_' + Date.now();
    const enc = new TextEncoder();

    const parts = [
      enc.encode(`--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n`),
      enc.encode(JSON.stringify(metadata)),
      enc.encode(`\r\n--${boundary}\r\nContent-Type: ${mime}\r\n\r\n`),
      uint8Data instanceof Uint8Array ? uint8Data : new Uint8Array(uint8Data),
      enc.encode(`\r\n--${boundary}--`)
    ];

    const totalLen = parts.reduce((s, p) => s + p.length, 0);
    const body = new Uint8Array(totalLen);
    let off = 0;
    for (const p of parts) { body.set(p, off); off += p.length; }

    const resp = await fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id',
      {
        method: 'POST',
        headers: {
          Authorization: 'Bearer ' + token,
          'Content-Type': `multipart/related; boundary=${boundary}`
        },
        body
      }
    );
    if (!resp.ok) throw new Error('Drive upload failed: ' + resp.statusText);
    const { id } = await resp.json();

    // Make publicly readable
    await fetch(`https://www.googleapis.com/drive/v3/files/${id}/permissions`, {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ role: 'reader', type: 'anyone' })
    });

    return {
      drive_file_id: id,
      drive_url: `https://drive.google.com/uc?export=view&id=${id}`,
      mime_type: mime
    };
  }

  // ── Rewrite media refs in HTML ─────────────────────────────────────────────
  function rewriteMediaRefs(html, mediaMap) {
    if (!html) return html;
    let out = html;
    out = out.replace(/<img([^>]*)\ssrc="([^"]+)"([^>]*)>/gi, (m, pre, src, post) => {
      const e = mediaMap[src];
      if (e?.drive_url) return `<img${pre} src="${e.drive_url}"${post}>`;
      return m;
    });
    out = out.replace(/\[sound:([^\]]+)\]/g, (m, fname) => {
      const e = mediaMap[fname];
      if (e?.drive_url) return `<audio controls src="${e.drive_url}"></audio>`;
      return m;
    });
    return out;
  }

  // ── Main run() ────────────────────────────────────────────────────────────────
  async function run({ file, deckFilter, mediaMode, token, onProgress }) {
    onProgress(2, 'Loading SQL engine…');
    const SQL = await ensureSQL();

    onProgress(5, 'Reading ZIP archive…');
    let zip;
    try {
      zip = await extractZip(file);
    } catch (e) {
      throw new Error('Could not open ZIP: ' + e.message);
    }

    // Locate the SQLite collection DB
    const dbName = ['collection.anki21b', 'collection.anki21', 'collection.anki2', 'collection']
      .find(n => zip.file(n));
    if (!dbName) throw new Error('No Anki collection database found in ZIP (expected collection.anki21 etc.).');

    onProgress(10, 'Opening Anki database…');
    const dbBuf = await zip.file(dbName).async('uint8array');
    const db = new SQL.Database(dbBuf);

    // ── Media manifest ────────────────────────────────────────────────────────
    // Anki stores a JSON map: { "0": "actual_filename.jpg", "1": "other.png" }
    // The ZIP entries use the numeric keys as filenames.
    const mediaManifest = {}; // numeric_key → real filename
    const mediaFile = zip.file('media');
    if (mediaFile) {
      try {
        const raw = await mediaFile.async('text');
        const manifest = JSON.parse(raw);
        for (const [k, v] of Object.entries(manifest)) mediaManifest[k] = v;
      } catch (e) {
        console.warn('Could not parse media manifest:', e);
      }
    }

    // ── Decks ─────────────────────────────────────────────────────────────────
    let deckIdToName = {};
    try {
      const res = db.exec('SELECT decks FROM col LIMIT 1');
      if (res.length) {
        const map = JSON.parse(res[0].values[0][0]);
        for (const [id, d] of Object.entries(map)) deckIdToName[id] = d.name || 'Default';
      }
    } catch (e) { console.warn('decks parse:', e); }

    // ── Note models ───────────────────────────────────────────────────────────
    let modelsMap = {};
    try {
      const res = db.exec('SELECT models FROM col LIMIT 1');
      if (res.length) modelsMap = JSON.parse(res[0].values[0][0]);
    } catch (e) { console.warn('models parse:', e); }

    // ── Notes ─────────────────────────────────────────────────────────────────
    onProgress(15, 'Reading notes…');
    let notesRes;
    try {
      notesRes = db.exec('SELECT id, guid, mid, tags, flds FROM notes');
    } catch (e) { throw new Error('Failed to read notes: ' + e.message); }
    if (!notesRes.length) throw new Error('No notes found in database.');

    const noteById = {};
    for (const [id, guid, mid, tags, flds] of notesRes[0].values) {
      noteById[String(id)] = { id: String(id), guid, mid: String(mid), tags, flds };
    }

    // ── Cards ─────────────────────────────────────────────────────────────────
    onProgress(20, 'Reading cards…');
    let cardsRes;
    try {
      cardsRes = db.exec('SELECT id, nid, did, ord, ivl, factor, reps, lapses, due, queue FROM cards');
    } catch (e) { throw new Error('Failed to read cards: ' + e.message); }
    if (!cardsRes.length) throw new Error('No cards found in database.');

    const allCards = cardsRes[0].values;
    const filtered = deckFilter
      ? allCards.filter(([,, did]) => {
          const n = deckIdToName[String(did)] || 'Default';
          return n.toLowerCase().startsWith(deckFilter.toLowerCase());
        })
      : allCards;

    onProgress(22, `Found ${filtered.length} card(s) to import.`);

    // ── Upload media ──────────────────────────────────────────────────────────
    const mediaMap = Object.assign({}, Media._map || {}); // local working copy
    let mediaUploaded = 0;

    if (mediaMode === 'drive' && Object.keys(mediaManifest).length > 0) {
      onProgress(25, 'Setting up Google Drive folder…');
      let folderId = null;
      try {
        folderId = await Media._ensureFolder(token);
      } catch (e) {
        console.warn('Could not create Drive folder:', e);
      }

      const entries = Object.entries(mediaManifest);
      const batchSave = [];

      for (let i = 0; i < entries.length; i++) {
        const [zipKey, filename] = entries[i];
        const pct = 25 + Math.round(((i + 1) / entries.length) * 35);
        onProgress(pct, `Uploading media (${i+1}/${entries.length}): ${filename}`);

        if (mediaMap[filename]?.drive_url) continue; // already uploaded

        const zipEntry = zip.file(zipKey);
        if (!zipEntry) continue;

        try {
          const data = await zipEntry.async('uint8array');
          const mime = mimeFor(filename);
          const result = await uploadToDrive(token, filename, data, mime, folderId);
          mediaMap[filename] = result;
          batchSave.push({ filename, ...result });
          mediaUploaded++;
        } catch (e) {
          console.warn(`Skip media ${filename}:`, e.message);
        }
      }

      if (batchSave.length > 0) {
        onProgress(60, `Saving ${batchSave.length} media record(s) to sheet…`);
        try {
          await Sheets.saveMediaBatch(batchSave);
          batchSave.forEach(e => { Media._map[e.filename] = e; });
        } catch (e) { console.warn('Media batch save failed:', e); }
      }
    }

    // ── Deduplicate against existing cards in sheet ───────────────────────────
    onProgress(62, 'Checking for duplicates…');
    const existingNoteIds = new Set((App.s.cards || []).map(c => c.note_id).filter(Boolean));

    // ── Build card + progress rows ────────────────────────────────────────────
    onProgress(65, 'Building card data…');

    const today = new Date().toISOString().split('T')[0];
    const cardBatch = [];
    const progBatch = [];

    for (let i = 0; i < filtered.length; i++) {
      const [cardId, nid, did, ord, ivl, factor, reps, lapses, due, queue] = filtered[i];
      const note = noteById[String(nid)];
      if (!note) continue;

      const noteIdStr = String(note.id);
      if (existingNoteIds.has(noteIdStr)) continue;

      const model = modelsMap[note.mid] || {};
      const { front, back } = renderNote(note.flds, model);
      if (!front.trim()) continue;

      const deckName = deckIdToName[String(did)] || 'Default';
      const ankiTags = (note.tags || '').trim().split(/\s+/).filter(Boolean).join(' :: ');

      const frontHtml = mediaMode === 'drive' ? rewriteMediaRefs(front, mediaMap) : front;
      const backHtml  = mediaMode === 'drive' ? rewriteMediaRefs(back,  mediaMap) : back;

      const newId = crypto.randomUUID();
      const createdAt = new Date().toISOString();

      cardBatch.push([
        newId, frontHtml, backHtml,
        ankiTags, createdAt, '',
        deckName, noteIdStr, model.name || '', String(lapses || 0)
      ]);

      // Convert Anki SM2: factor stored as 10× (2500 → 2.5), ivl in days
      const easeFactor = ((factor || 2500) / 1000).toFixed(4);
      const interval   = Math.max(0, ivl || 0);
      let dueDate = today;
      if (queue === 2 && ivl > 0) {
        const colCreation = (() => {
          try { return db.exec('SELECT crt FROM col LIMIT 1')[0].values[0][0] * 1000; } catch (_) { return Date.now(); }
        })();
        const dueDateMs = colCreation + due * 86400000;
        dueDate = new Date(dueDateMs).toISOString().split('T')[0];
      }

      progBatch.push([
        newId,
        String(easeFactor), String(interval), String(Math.max(0, reps || 0)),
        dueDate, '', String(lapses || 0)
      ]);

      existingNoteIds.add(noteIdStr);

      if ((i + 1) % 100 === 0) {
        const pct = 65 + Math.round(((i + 1) / filtered.length) * 15);
        onProgress(pct, `Prepared ${cardBatch.length} card(s)…`);
      }
    }

    // ── Write to sheet in chunks ──────────────────────────────────────────────
    const CHUNK = 500;
    let written = 0;

    for (let i = 0; i < cardBatch.length; i += CHUNK) {
      const pct = 80 + Math.round((i / Math.max(cardBatch.length, 1)) * 15);
      onProgress(pct, `Writing cards ${i+1}–${Math.min(i+CHUNK, cardBatch.length)} to sheet…`);
      await Sheets.appendBatch(Sheets.CARDS, cardBatch.slice(i, i + CHUNK));
      written += Math.min(CHUNK, cardBatch.length - i);
    }

    for (let i = 0; i < progBatch.length; i += CHUNK) {
      await Sheets.appendBatch(Sheets.PROGRESS, progBatch.slice(i, i + CHUNK));
    }

    onProgress(99, `Done — ${written} card(s) written.`);
    db.close();
    return { cards: written, mediaUploaded };
  }

  return { run };
})();
