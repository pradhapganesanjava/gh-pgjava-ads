// Google Sheets API wrapper
// Sheet "Cards":     id | front | back | tags | created_at | notes | deck | note_id | model_name | lapses
// Sheet "Progress":  card_id | ease_factor | interval | repetitions | due_date | last_review | lapses
// Sheet "Media":     id | filename | drive_file_id | drive_url | mime_type | created_at
// Sheet "TagGroups": id | name | tags | created_at
// Sheet "Settings":  key | value

const Sheets = {
  CARDS:     'Cards',
  PROGRESS:  'Progress',
  MEDIA:     'Media',
  TAGGROUPS: 'TagGroups',
  SETTINGS:  'Settings',
  CARDS_HDR:     ['id','front','back','tags','created_at','notes','deck','note_id','model_name','lapses'],
  PROGRESS_HDR:  ['card_id','ease_factor','interval','repetitions','due_date','last_review','lapses'],
  MEDIA_HDR:     ['id','filename','drive_file_id','drive_url','mime_type','created_at'],
  TAGGROUPS_HDR: ['id','name','tags','created_at'],
  SETTINGS_HDR:  ['key','value'],
  _ready: false,

  // ── Init ──────────────────────────────────────────────────────────
  init() {
    if (this._ready) return Promise.resolve();
    return new Promise((resolve, reject) => {
      gapi.load('client', async () => {
        try {
          const opts = { discoveryDocs: [
            'https://sheets.googleapis.com/$discovery/rest?version=v4',
            'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest'
          ] };
          if (Config.googleApiKey) opts.apiKey = Config.googleApiKey;
          await gapi.client.init(opts);
          this._ready = true;
          resolve();
        } catch (e) { reject(e); }
      });
    });
  },

  setToken(token) {
    gapi.client.setToken({ access_token: token });
  },

  // Extract a readable message from gapi error objects or standard Errors
  _errMsg(e) {
    if (!e) return 'Unknown error';
    if (e.result?.error?.message) return e.result.error.message;
    if (e.message) return e.message;
    if (typeof e === 'string') return e;
    return JSON.stringify(e);
  },

  // ── Low-level helpers ─────────────────────────────────────────────
  async get(range) {
    try {
      const r = await gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: Config.sheetId, range
      });
      return r.result.values || [];
    } catch (e) {
      const msg = this._errMsg(e);
      if (msg.includes('Unable to parse range') || e.status === 400) return [];
      throw new Error(msg);
    }
  },

  async set(range, values) {
    try {
      await gapi.client.sheets.spreadsheets.values.update({
        spreadsheetId: Config.sheetId, range,
        valueInputOption: 'RAW',
        resource: { values }
      });
    } catch (e) { throw new Error(this._errMsg(e)); }
  },

  async append(sheet, values) {
    try {
      await gapi.client.sheets.spreadsheets.values.append({
        spreadsheetId: Config.sheetId,
        range: `${sheet}!A1`,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: { values }
      });
    } catch (e) { throw new Error(this._errMsg(e)); }
  },

  // Batch append many rows at once (more efficient for imports)
  async appendBatch(sheet, rows) {
    if (!rows.length) return;
    try {
      await gapi.client.sheets.spreadsheets.values.append({
        spreadsheetId: Config.sheetId,
        range: `${sheet}!A1`,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        resource: { values: rows }
      });
    } catch (e) { throw new Error(this._errMsg(e)); }
  },

  // Ensure header rows exist in all sheets
  async ensureHeaders() {
    const specs = [
      [this.CARDS,     this.CARDS_HDR,     'A1:J1'],
      [this.PROGRESS,  this.PROGRESS_HDR,  'A1:G1'],
      [this.MEDIA,     this.MEDIA_HDR,     'A1:F1'],
      [this.TAGGROUPS, this.TAGGROUPS_HDR, 'A1:D1'],
      [this.SETTINGS,  this.SETTINGS_HDR,  'A1:B1']
    ];
    for (const [sheet, hdr, range] of specs) {
      try {
        const rows = await this.get(`${sheet}!${range}`);
        if (!rows.length) await this.set(`${sheet}!${range}`, [hdr]);
      } catch (e) {
        console.warn(`Sheet tab "${sheet}" not found:`, this._errMsg(e));
      }
    }
  },

  // ── Sheet discovery ───────────────────────────────────────────────
  async findOrCreateSheet(name = 'PGAnkiADS') {
    try {
      const res = await gapi.client.drive.files.list({
        q: `name='${name}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`,
        fields: 'files(id,name)',
        spaces: 'drive',
        pageSize: 1
      });
      const files = (res.result.files || []);
      if (files.length > 0) return files[0].id;
    } catch (e) {
      console.warn('Drive search failed — will create new sheet:', this._errMsg(e));
    }

    const created = await gapi.client.sheets.spreadsheets.create({
      resource: { properties: { title: name } }
    });
    return created.result.spreadsheetId;
  },

  async checkAccess() {
    try {
      await gapi.client.sheets.spreadsheets.get({ spreadsheetId: Config.sheetId });
    } catch (e) {
      const msg = this._errMsg(e);
      if (e.status === 404) throw new Error('Sheet not found — check your Sheet ID in Settings.');
      if (e.status === 403) throw new Error('No access to this Sheet — make sure it is shared with your Google account.');
      throw new Error('Sheet error: ' + msg);
    }
  },

  // ── Cards CRUD ────────────────────────────────────────────────────
  async loadCards() {
    try {
      const rows = await this.get(`${this.CARDS}!A2:J`);
      return rows
        .map(r => ({
          id:         r[0],
          front:      r[1],
          back:       r[2],
          tags:       r[3]||'',
          created_at: r[4]||'',
          notes:      r[5]||'',
          deck:       r[6]||'Default',
          note_id:    r[7]||'',
          model_name: r[8]||'',
          lapses:     parseInt(r[9])||0
        }))
        .filter(c => c.id && c.front);
    } catch (e) { throw new Error('Loading cards failed: ' + this._errMsg(e)); }
  },

  async saveCard(card) {
    card.id = card.id || crypto.randomUUID();
    card.created_at = card.created_at || new Date().toISOString();
    await this.append(this.CARDS, [[
      card.id, card.front, card.back,
      card.tags||'', card.created_at, card.notes||'',
      card.deck||'Default', card.note_id||'', card.model_name||'',
      String(card.lapses||0)
    ]]);
    return card;
  },

  async updateCard(card) {
    const rows = await this.get(`${this.CARDS}!A:A`);
    const idx = rows.findIndex(r => r[0] === card.id);
    if (idx < 1) throw new Error('Card not found in sheet.');
    await this.set(`${this.CARDS}!A${idx+1}:J${idx+1}`,
      [[card.id, card.front, card.back, card.tags||'', card.created_at||'', card.notes||'',
        card.deck||'Default', card.note_id||'', card.model_name||'', String(card.lapses||0)]]);
  },

  async deleteCard(cardId) {
    const rows = await this.get(`${this.CARDS}!A:A`);
    const idx = rows.findIndex(r => r[0] === cardId);
    if (idx < 1) return;
    await this.set(`${this.CARDS}!A${idx+1}:J${idx+1}`, [['','','','','','','','','','']]);
  },

  // ── Progress CRUD ─────────────────────────────────────────────────
  async loadProgress() {
    try {
      const rows = await this.get(`${this.PROGRESS}!A2:G`);
      const map = {};
      rows.forEach(r => {
        if (!r[0]) return;
        map[r[0]] = {
          card_id:     r[0],
          easeFactor:  parseFloat(r[1]) || 2.5,
          interval:    parseInt(r[2])   || 0,
          repetitions: parseInt(r[3])   || 0,
          dueDate:     r[4] || new Date().toISOString().split('T')[0],
          lastReview:  r[5] || '',
          lapses:      parseInt(r[6])   || 0
        };
      });
      return map;
    } catch (e) { throw new Error('Loading progress failed: ' + this._errMsg(e)); }
  },

  async saveProgress(p) {
    const rows = await this.get(`${this.PROGRESS}!A:A`);
    const row = [p.card_id, String(p.easeFactor), String(p.interval), String(p.repetitions),
                 p.dueDate, p.lastReview||'', String(p.lapses||0)];
    const idx = rows.findIndex(r => r[0] === p.card_id);
    if (idx < 1) {
      await this.append(this.PROGRESS, [row]);
    } else {
      await this.set(`${this.PROGRESS}!A${idx+1}:G${idx+1}`, [row]);
    }
  },

  // ── Media CRUD ────────────────────────────────────────────────────
  async loadMedia() {
    try {
      const rows = await this.get(`${this.MEDIA}!A2:F`);
      const map = {};
      rows.forEach(r => {
        if (!r[0]) return;
        const entry = { id: r[0], filename: r[1], drive_file_id: r[2], drive_url: r[3], mime_type: r[4]||'', created_at: r[5]||'' };
        map[entry.filename] = entry;
      });
      return map; // keyed by filename
    } catch (e) { return {}; }
  },

  async saveMediaEntry(entry) {
    entry.id = entry.id || crypto.randomUUID();
    entry.created_at = entry.created_at || new Date().toISOString();
    // Check if filename already exists, update if so
    const rows = await this.get(`${this.MEDIA}!B:B`);
    const idx = rows.findIndex(r => r[0] === entry.filename);
    const row = [entry.id, entry.filename, entry.drive_file_id||'', entry.drive_url||'', entry.mime_type||'', entry.created_at];
    if (idx < 1) {
      await this.append(this.MEDIA, [row]);
    } else {
      // Find actual row index in sheet (B column offset = same row in A column)
      const aRows = await this.get(`${this.MEDIA}!A:A`);
      const aIdx = aRows.findIndex((r, i) => i > 0 && r[0]);
      // Simple approach: re-check via ID column
      const allRows = await this.get(`${this.MEDIA}!A2:F`);
      const rowIdx = allRows.findIndex(r => r[1] === entry.filename);
      if (rowIdx >= 0) {
        await this.set(`${this.MEDIA}!A${rowIdx+2}:F${rowIdx+2}`, [row]);
      } else {
        await this.append(this.MEDIA, [row]);
      }
    }
    return entry;
  },

  // Batch save media entries (for import)
  async saveMediaBatch(entries) {
    if (!entries.length) return;
    const rows = entries.map(e => [
      e.id || crypto.randomUUID(),
      e.filename, e.drive_file_id||'', e.drive_url||'',
      e.mime_type||'', e.created_at || new Date().toISOString()
    ]);
    await this.appendBatch(this.MEDIA, rows);
  },

  // ── TagGroups CRUD ────────────────────────────────────────────────
  async loadTagGroups() {
    try {
      const rows = await this.get(`${this.TAGGROUPS}!A2:D`);
      return rows
        .map(r => ({ id: r[0], name: r[1], tags: r[2]||'', created_at: r[3]||'' }))
        .filter(g => g.id && g.name);
    } catch (e) { return []; }
  },

  async saveTagGroup(group) {
    group.id = group.id || crypto.randomUUID();
    group.created_at = group.created_at || new Date().toISOString();
    await this.append(this.TAGGROUPS, [[group.id, group.name, group.tags||'', group.created_at]]);
    return group;
  },

  async deleteTagGroup(groupId) {
    const rows = await this.get(`${this.TAGGROUPS}!A:A`);
    const idx = rows.findIndex(r => r[0] === groupId);
    if (idx < 1) return;
    await this.set(`${this.TAGGROUPS}!A${idx+1}:D${idx+1}`, [['','','','']]);
  },

  // ── Settings CRUD ─────────────────────────────────────────────────
  async loadSettings() {
    try {
      const rows = await this.get(`${this.SETTINGS}!A2:B`);
      const map = {};
      rows.forEach(r => { if (r[0]) map[r[0]] = r[1] || ''; });
      return map;
    } catch (e) { return {}; }
  },

  async saveSetting(key, value) {
    const rows = await this.get(`${this.SETTINGS}!A:A`);
    const idx = rows.findIndex(r => r[0] === key);
    if (idx < 1) {
      await this.append(this.SETTINGS, [[key, value]]);
    } else {
      await this.set(`${this.SETTINGS}!A${idx+1}:B${idx+1}`, [[key, value]]);
    }
  }
};
