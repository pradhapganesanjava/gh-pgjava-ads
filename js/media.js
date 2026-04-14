// Google Drive media manager for Anki media files
// Handles uploading media to Drive, tracking filenames → Drive URLs,
// and processing card HTML to replace local media refs with Drive URLs.
const Media = {
  _token: null,
  _map: null,          // filename → { drive_file_id, drive_url, mime_type }
  _folderId: null,     // cached Drive folder ID

  setToken(token) {
    this._token = token;
  },

  // Load media map from Sheets (call once after sign-in)
  async init() {
    this._map = await Sheets.loadMedia();
  },

  // Alias used by both App._load() and AnkiImporter
  async loadMap() {
    this._map = await Sheets.loadMedia();
    return this._map;
  },

  // Return in-memory media map (load first if not loaded)
  async getMap() {
    if (!this._map) this._map = await Sheets.loadMedia();
    return this._map;
  },

  // Alias for ensureFolder — used by AnkiImporter which passes its own token
  async _ensureFolder(token) {
    const savedToken = this._token;
    this._token = token;
    const id = await this.ensureFolder();
    this._token = savedToken;
    return id;
  },

  // Replace Anki media filenames in HTML with Drive URLs
  // e.g. <img src="paste-xxx.png"> → <img src="https://drive.google.com/uc?export=view&id=FILE_ID">
  async processHtml(html) {
    if (!html) return html;
    const map = await this.getMap();
    // Replace img src, audio src, video src that are plain filenames (not URLs)
    return html.replace(/(<(?:img|audio|video|source)[^>]*\s(?:src|data-src)=")([^"http][^"]*?)(")/gi,
      (match, pre, filename, post) => {
        const entry = map[filename];
        if (entry?.drive_url) return `${pre}${entry.drive_url}${post}`;
        return match;
      }
    );
  },

  // ── Google Drive operations ───────────────────────────────────────

  // Ensure a Drive folder exists for Anki media; returns folder ID
  async ensureFolder(folderName = 'PGAnkiADS-Media') {
    if (this._folderId) return this._folderId;
    if (Config.mediaDriveFolderId) {
      this._folderId = Config.mediaDriveFolderId;
      return this._folderId;
    }

    // Search for existing folder
    const searchRes = await fetch(
      `https://www.googleapis.com/drive/v3/files?q=name%3D'${encodeURIComponent(folderName)}'%20and%20mimeType%3D'application%2Fvnd.google-apps.folder'%20and%20trashed%3Dfalse&fields=files(id,name)`,
      { headers: { Authorization: `Bearer ${this._token}` } }
    );
    const searchData = await searchRes.json();
    if (searchData.files?.length > 0) {
      this._folderId = searchData.files[0].id;
      Config.mediaDriveFolderId = this._folderId;
      return this._folderId;
    }

    // Create new folder
    const createRes = await fetch('https://www.googleapis.com/drive/v3/files', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${this._token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        name: folderName,
        mimeType: 'application/vnd.google-apps.folder'
      })
    });
    const folderData = await createRes.json();
    this._folderId = folderData.id;
    Config.mediaDriveFolderId = this._folderId;

    // Make folder publicly accessible (anyone with link can view)
    await this._makePublic(this._folderId);

    return this._folderId;
  },

  // Upload a file blob to Google Drive; returns { file_id, drive_url }
  async uploadFile(blob, filename, folderId) {
    const mime = blob.type || this._guessMime(filename);

    // Multipart upload: metadata + file content
    const boundary = '-------314159265358979323846';
    const metadata = JSON.stringify({ name: filename, parents: [folderId] });
    const body = [
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${metadata}\r\n`,
      `--${boundary}\r\nContent-Type: ${mime}\r\n\r\n`,
      blob,
      `\r\n--${boundary}--`
    ];

    const formData = new Blob(body, { type: `multipart/related; boundary="${boundary}"` });

    const res = await fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webContentLink',
      {
        method: 'POST',
        headers: { Authorization: `Bearer ${this._token}` },
        body: formData
      }
    );
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(`Drive upload failed: ${err.error?.message || res.status}`);
    }
    const data = await res.json();
    const fileId = data.id;

    // Make file publicly viewable
    await this._makePublic(fileId);

    const driveUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;
    return { file_id: fileId, drive_url: driveUrl };
  },

  // Make a Drive file/folder publicly readable (anyone with link)
  async _makePublic(fileId) {
    await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}/permissions`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${this._token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ role: 'reader', type: 'anyone' })
    });
  },

  _guessMime(filename) {
    const ext = (filename.split('.').pop() || '').toLowerCase();
    const mimes = {
      jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png',
      gif: 'image/gif', svg: 'image/svg+xml', webp: 'image/webp',
      mp3: 'audio/mpeg', ogg: 'audio/ogg', wav: 'audio/wav',
      mp4: 'video/mp4', webm: 'video/webm',
      pdf: 'application/pdf'
    };
    return mimes[ext] || 'application/octet-stream';
  },

  // Update in-memory map and save to Sheets
  async registerMedia(entry) {
    if (!this._map) this._map = {};
    this._map[entry.filename] = entry;
    await Sheets.saveMediaEntry(entry);
  },

  // Register many at once (for batch import)
  async registerMediaBatch(entries) {
    if (!this._map) this._map = {};
    for (const e of entries) this._map[e.filename] = e;
    await Sheets.saveMediaBatch(entries);
  },

  // ── Anki .colpkg import ───────────────────────────────────────────

  // Parse and import an Anki .colpkg file.
  // onProgress(msg, pct) is called with status updates.
  // Returns { cardsImported, mediaUploaded, errors }
  async importColpkg(file, onProgress = () => {}) {
    const result = { cardsImported: 0, mediaUploaded: 0, errors: [] };

    onProgress('Extracting .colpkg archive…', 5);
    let zip;
    try {
      zip = await JSZip.loadAsync(file);
    } catch (e) {
      throw new Error('Could not read .colpkg file: ' + e.message);
    }

    // Find the SQLite database file
    const dbFile = zip.file('collection.anki21') || zip.file('collection.anki2') || zip.file('collection.anki21b');
    if (!dbFile) throw new Error('No Anki collection database found in archive.');

    onProgress('Loading SQLite database…', 10);
    const dbBuffer = await dbFile.async('arraybuffer');

    // Initialize sql.js with WASM
    let SQL;
    try {
      SQL = await initSqlJs({ locateFile: f => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/${f}` });
    } catch (e) {
      throw new Error('Could not load sql.js WASM: ' + e.message);
    }
    const db = new SQL.Database(new Uint8Array(dbBuffer));

    // Parse media manifest (maps numbered files to original filenames)
    onProgress('Parsing media manifest…', 15);
    const mediaFile = zip.file('media');
    let mediaManifest = {};
    if (mediaFile) {
      try {
        const raw = await mediaFile.async('string');
        mediaManifest = JSON.parse(raw);
      } catch (e) {
        console.warn('Could not parse media manifest:', e.message);
      }
    }
    // Invert: { "0": "image.png" } → { "image.png": "0" }
    const filenameToNumber = {};
    for (const [num, fname] of Object.entries(mediaManifest)) {
      filenameToNumber[fname] = num;
    }

    // Parse col table for decks and models
    onProgress('Parsing deck structure…', 20);
    const colRow = db.exec('SELECT decks, models FROM col LIMIT 1');
    let decks = {}, models = {};
    if (colRow.length > 0 && colRow[0].values.length > 0) {
      try { decks = JSON.parse(colRow[0].values[0][0]); } catch {}
      try { models = JSON.parse(colRow[0].values[0][1]); } catch {}
    }

    // Build deck id → full name map
    const deckNames = {};
    for (const [id, deck] of Object.entries(decks)) {
      if (deck.name !== '__cram_deck__') deckNames[id] = deck.name;
    }

    // Build model id → {fields, templates} map
    const modelInfo = {};
    for (const [id, model] of Object.entries(models)) {
      modelInfo[id] = {
        name: model.name || 'Basic',
        fields: (model.flds || []).map(f => f.name),
        templates: model.tmpls || []
      };
    }

    // Collect all media filenames referenced in cards
    onProgress('Scanning card content for media…', 25);
    const notesResult = db.exec('SELECT id, guid, mid, tags, flds FROM notes');
    const allNotes = notesResult.length > 0 ? notesResult[0].values : [];

    const referencedMedia = new Set();
    for (const [,,,, flds] of allNotes) {
      const matches = flds.matchAll(/<(?:img|audio|source)[^>]*\s(?:src|data-src)="([^"]+)"/gi);
      for (const m of matches) referencedMedia.add(m[1]);
      // Also match [sound:filename] Anki format
      const soundMatches = flds.matchAll(/\[sound:([^\]]+)\]/g);
      for (const m of soundMatches) referencedMedia.add(m[1]);
    }

    // Upload referenced media files to Drive
    const totalMedia = referencedMedia.size;
    if (totalMedia > 0) {
      onProgress(`Uploading ${totalMedia} media file(s) to Google Drive…`, 30);
      const folderId = await this.ensureFolder();
      const existingMap = await this.getMap();
      const toUpload = [...referencedMedia].filter(f => !existingMap[f]);
      let uploaded = 0;
      const newEntries = [];

      for (const filename of toUpload) {
        const zipNum = filenameToNumber[filename];
        if (!zipNum) continue;
        const zipEntry = zip.file(zipNum);
        if (!zipEntry) continue;

        try {
          const blob = new Blob([await zipEntry.async('arraybuffer')], { type: this._guessMime(filename) });
          const { file_id, drive_url } = await this.uploadFile(blob, filename, folderId);
          newEntries.push({ filename, drive_file_id: file_id, drive_url, mime_type: blob.type });
          uploaded++;
          const pct = 30 + Math.round((uploaded / toUpload.length) * 35);
          onProgress(`Uploading media ${uploaded}/${toUpload.length}: ${filename}`, pct);
        } catch (e) {
          result.errors.push(`Media upload failed for ${filename}: ${e.message}`);
        }
      }

      if (newEntries.length) {
        await this.registerMediaBatch(newEntries);
        result.mediaUploaded = newEntries.length;
      }
    }

    // Now import cards
    onProgress('Reading cards from Anki database…', 65);
    const cardsResult = db.exec('SELECT nid, did, ivl, factor, reps, lapses, due, type FROM cards WHERE queue >= 0');
    const ankiCards = cardsResult.length > 0 ? cardsResult[0].values : [];

    // Group cards by note_id (pick first card per note for simplicity)
    const noteCards = {};
    for (const [nid, did, ivl, factor, reps, lapses, due, type] of ankiCards) {
      if (!noteCards[nid]) {
        noteCards[nid] = { did, ivl, factor, reps, lapses, due, type };
      }
    }

    // Build the cards to import
    const mediaMap = await this.getMap();
    const cardsToInsert = [];
    const progressToInsert = [];
    let noteIdx = 0;

    onProgress('Processing notes…', 70);
    for (const [noteIdStr, , midStr, tagsStr, fldsStr] of allNotes) {
      const noteId = String(noteIdStr);
      const cardInfo = noteCards[noteId];
      if (!cardInfo) continue; // skip notes with no active cards

      const mid = String(midStr);
      const model = modelInfo[mid] || { name: 'Basic', fields: ['Front', 'Back'], templates: [] };
      const fields = fldsStr.split('\x1f');

      // Get front and back from the model's first template
      let front = fields[0] || '';
      let back  = fields[1] || '';

      // Process template-based rendering if template exists
      if (model.templates.length > 0) {
        const tmpl = model.templates[0];
        front = this._renderAnkiTemplate(tmpl.qfmt || '{{Front}}', fields, model.fields);
        back  = this._renderAnkiTemplate(tmpl.afmt || '{{Back}}', fields, model.fields);
        // Remove the front content from back if it uses {{FrontSide}}
        back = back.replace(/{{FrontSide}}/gi, '');
      }

      // Replace media refs with Drive URLs in card content
      front = this._replaceMediaRefs(front, mediaMap);
      back  = this._replaceMediaRefs(back, mediaMap);

      // Convert [sound:file] to HTML audio
      front = front.replace(/\[sound:([^\]]+)\]/g, (m, f) => {
        const entry = mediaMap[f];
        const src = entry?.drive_url || f;
        return `<audio controls src="${src}" style="max-width:100%"></audio>`;
      });
      back = back.replace(/\[sound:([^\]]+)\]/g, (m, f) => {
        const entry = mediaMap[f];
        const src = entry?.drive_url || f;
        return `<audio controls src="${src}" style="max-width:100%"></audio>`;
      });

      // Parse Anki tags (space-separated in Anki, convert to :: hierarchy format)
      const ankiTags = (tagsStr || '').trim().split(/\s+/).filter(Boolean);
      const tags = ankiTags.map(t => t.replace(/::/g, '::')).join(' :: ');

      // Get deck name
      const deckPath = (deckNames[String(cardInfo.did)] || 'Default').replace(/::/g, '::');

      const cardId = crypto.randomUUID();
      cardsToInsert.push([
        cardId, front.trim(), back.trim(), tags,
        new Date().toISOString(), '', deckPath,
        noteId, model.name, String(cardInfo.lapses || 0)
      ]);

      // Import progress from Anki SM data
      // Anki ivl: positive = days, negative = seconds (learning)
      const interval = Math.max(0, cardInfo.ivl || 0);
      const easeFactor = Math.max(1.3, (cardInfo.factor || 2500) / 1000);
      const repetitions = cardInfo.reps || 0;
      const lapses = cardInfo.lapses || 0;
      // Calculate due date from Anki's due field
      const dueDate = this._ankiDueToDate(cardInfo.due, cardInfo.type, interval);
      progressToInsert.push([
        cardId, String(easeFactor), String(interval),
        String(repetitions), dueDate, '', String(lapses)
      ]);

      noteIdx++;
      if (noteIdx % 100 === 0) {
        const pct = 70 + Math.round((noteIdx / allNotes.length) * 20);
        onProgress(`Processing card ${noteIdx}/${allNotes.length}…`, pct);
      }
    }

    // Batch insert cards to Google Sheets
    onProgress(`Saving ${cardsToInsert.length} cards to Google Sheets…`, 90);
    const BATCH = 500; // Google Sheets API limit
    for (let i = 0; i < cardsToInsert.length; i += BATCH) {
      await Sheets.appendBatch(Sheets.CARDS, cardsToInsert.slice(i, i + BATCH));
      await Sheets.appendBatch(Sheets.PROGRESS, progressToInsert.slice(i, i + BATCH));
      const pct = 90 + Math.round((i / cardsToInsert.length) * 9);
      onProgress(`Saving batch ${Math.floor(i/BATCH)+1}…`, pct);
    }

    db.close();
    result.cardsImported = cardsToInsert.length;
    onProgress('Import complete!', 100);
    return result;
  },

  // Render Anki template (basic field substitution)
  _renderAnkiTemplate(template, fields, fieldNames) {
    let result = template;
    for (let i = 0; i < fieldNames.length; i++) {
      const name = fieldNames[i];
      const value = fields[i] || '';
      // Replace {{FieldName}} and {{text:FieldName}} variants
      result = result.replace(new RegExp(`{{(?:text:)?${this._escapeRegex(name)}}}`, 'gi'), value);
      result = result.replace(new RegExp(`{{type:${this._escapeRegex(name)}}}`, 'gi'), `<input type="text" class="anki-type-input" placeholder="Type ${name}…">`);
    }
    // Remove unfilled template placeholders
    result = result.replace(/{{[^}]+}}/g, '');
    // Handle cloze deletions: {{c1::answer}} → answer (reveal form)
    result = result.replace(/{{c\d+::([^:}]+)(?:::[^}]+)?}}/g, '<span class="cloze">$1</span>');
    return result;
  },

  _escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  },

  _replaceMediaRefs(html, mediaMap) {
    return html.replace(/(<(?:img|audio|video|source)[^>]*\s(?:src|data-src)=")([^"]+?)(")/gi,
      (match, pre, filename, post) => {
        if (filename.startsWith('http') || filename.startsWith('data:')) return match;
        const entry = mediaMap[filename];
        if (entry?.drive_url) return `${pre}${entry.drive_url}${post}`;
        return match;
      }
    );
  },

  // Convert Anki due field to ISO date string
  // type: 0=learning, 1=review, 2=relearn, 3=cram
  _ankiDueToDate(due, type, interval) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    if (type === 1 && due > 0) {
      // Review card: due is a day number from Anki epoch (Jan 1, 2006)
      const ankiEpoch = new Date(2006, 0, 1).getTime();
      const dueMs = ankiEpoch + due * 86400000;
      const dueDate = new Date(dueMs);
      return dueDate.toISOString().split('T')[0];
    }
    // Learning/new: calculate from interval
    if (interval > 0) {
      const d = new Date(today);
      d.setDate(d.getDate() + interval);
      return d.toISOString().split('T')[0];
    }
    // New card: due today
    return today.toISOString().split('T')[0];
  }
};
