// ─── Azure TTS voices ─────────────────────────────────────────────────────────
const TTS_VOICES = [
  { id: 'nova',    label: 'Nova',    desc: 'Female · warm' },
  { id: 'shimmer', label: 'Shimmer', desc: 'Female · soft' },
  { id: 'alloy',   label: 'Alloy',   desc: 'Neutral · balanced' },
  { id: 'echo',    label: 'Echo',    desc: 'Male · clear' },
  { id: 'fable',   label: 'Fable',   desc: 'British · warm' },
  { id: 'onyx',    label: 'Onyx',    desc: 'Male · deep' },
];

// ─── Theme definitions ────────────────────────────────────────────────────────
const THEMES = [
  { id: 'dark',     label: 'Dark',     bg: '#0f0f13', accent: '#6366f1', dot: '#6366f1' },
  { id: 'light',    label: 'Light',    bg: '#f5f5ff', accent: '#6366f1', dot: '#c8c8e8' },
  { id: 'soft',     label: 'Soft',     bg: '#1e1b34', accent: '#a78bfa', dot: '#a78bfa' },
  { id: 'contrast', label: 'Contrast', bg: '#000000', accent: '#faff00', dot: '#555555' },
  { id: 'glow',     label: 'Glow',     bg: '#050510', accent: '#00e5ff', dot: '#00e5ff' },
  { id: 'cartoon',  label: 'Cartoon',  bg: '#fff9e6', accent: '#7c3aed', dot: '#f0c040' },
];

// ─── App — main controller ────────────────────────────────────────────────────
const App = {
  s: {  // state
    view: 'loading',
    cards: [],
    progress: {},
    queue: [],
    qIdx: 0,
    flipped: false,
    aiEvalText: '',
    aiEvalLoading: false,
    sessionReviewed: 0,
    sessionCorrect: 0,
    editCard: null,
    searchQ: '',
    audioOn: localStorage.getItem('pgads_audio') !== 'false',  // default on
    userAnswer: '',
    isListening: false,
    selectedTags: [],
    tagSearch: '',
    tagGroups: [],
    answerExpanded: false,
    answerVisible: false,
    answerColCollapsed: false,
    browseFilterTags: [],
    browseSelectedCard: null,
    browseTagSearch: '',
    browseSearch: '',
    tagTreeExpanded: {},
    reviewTreeExpanded: {},
    tagsDrawerOpen: false,
    browseTagsDrawerOpen: false,
    // Decks
    selectedDeck: [],
    reviewLeftTab: 'tags',       // 'tags' | 'decks'
    reviewDeckSearch: '',
    reviewDeckTreeExpanded: {},
    browseDeckFilter: [],
    browseLeftTab: 'tags',       // 'tags' | 'decks'
    browseDeckSearch: '',
    deckTreeExpanded: {},
    browseSelectedCards: [],     // multi-select card IDs
    browseLeftCollapsed: true,
    reviewLeftCollapsed: true,
    deckManagerExpanded: {},
    decksShowNewForm: false,
    decksNewInput: '',
  },

  // ── Bootstrap ───────────────────────────────────────────────────────────────
  async init() {
    // Apply saved theme immediately before any render
    this._applyTheme(Config.theme, false);

    // Single global keydown listener
    document.addEventListener('keydown', e => this._onKey(e));

    if (!Config.isConfigured()) { this._render('setup'); return; }

    try {
      const restored = await Auth.init();
      if (restored) {
        await this._load();
        this._render('review');
      } else {
        this._render('login');
      }
    } catch (e) {
      console.error(e);
      if (Auth.isSignedIn()) {
        this._toast(e.message, 'error');
        this._render('settings');
      } else {
        this._render('login');
      }
    }
  },

  // Show a status message in the loading screen
  _setLoadMsg(msg) {
    const el = document.querySelector('#app p');
    if (el) el.textContent = msg;
  },

  async _load() {
    await Sheets.init();
    Sheets.setToken(Auth.token);

    // Validate the sheet is accessible (ID from ENV secret or localStorage)
    await Sheets.checkAccess();

    await Sheets.ensureHeaders();

    // Load settings from sheet — apply API keys, theme, etc.
    const settings = await Sheets.loadSettings();
    this._applySheetSettings(settings);

    [this.s.cards, this.s.progress, this.s.tagGroups] = await Promise.all([
      Sheets.loadCards(),
      Sheets.loadProgress(),
      Sheets.loadTagGroups()
    ]);
    // Init Media module with token + cached media map
    Media.setToken(Auth.token);
    await Media.loadMap();
    this._buildQueue();
  },

  // Apply key-value settings loaded from the Settings sheet
  _applySheetSettings(settings) {
    // Theme: sheet wins; no need to write back (we just read it)
    if (settings.theme) this._applyTheme(settings.theme, false);

    // Azure API key — bidirectional sync:
    //   sheet → localStorage  (normal case: roaming to a new device)
    //   localStorage → sheet  (migration: key was in old config, sheet is blank)
    if (settings.azureApiKey) {
      Config.azureApiKey = settings.azureApiKey;
    } else if (Config.azureApiKey) {
      Sheets.saveSetting('azureApiKey', Config.azureApiKey).catch(() => {});
    }

    if (settings.azureEndpoint)   Config.azureEndpoint   = settings.azureEndpoint;
    if (settings.azureDeployment) Config.azureDeployment = settings.azureDeployment;
    if (settings.azureApiVersion) Config.azureApiVersion = settings.azureApiVersion;

    // TTS-specific overrides (separate resource)
    if (settings.azureTtsEndpoint)   Config.azureTtsEndpoint   = settings.azureTtsEndpoint;
    if (settings.azureTtsApiKey)     Config.azureTtsApiKey     = settings.azureTtsApiKey;
    if (settings.azureTtsDeployment) Config.azureTtsDeployment = settings.azureTtsDeployment;
    if (settings.azureTtsApiVersion) Config.azureTtsApiVersion = settings.azureTtsApiVersion;
    if (settings.ttsVoice)           Config.ttsVoice           = settings.ttsVoice;
  },

  // Apply a theme: set data-theme on <html> and persist to localStorage
  // saveToSheet: also upsert the theme key in the Settings sheet
  _applyTheme(themeId, saveToSheet = false) {
    const id = THEMES.find(t => t.id === themeId) ? themeId : 'dark';
    document.documentElement.setAttribute('data-theme', id);
    Config.theme = id;
    if (saveToSheet && Auth.isSignedIn()) {
      Sheets.saveSetting('theme', id).catch(() => {});
    }
  },

  _buildQueue() {
    let cards = this.s.cards.filter(c => SM2.isDue(this.s.progress[c.id]));
    if (this.s.selectedTags.length > 0) {
      cards = cards.filter(c => {
        const ct = parseTags(c.tags || '');
        return this.s.selectedTags.some(sel => ct.some(t => t === sel || t.startsWith(sel + '::')));
      });
    }
    if (this.s.selectedDeck.length > 0) {
      cards = cards.filter(c => {
        const d = (c.deck || 'Default').trim();
        return this.s.selectedDeck.some(sel => d === sel || d.startsWith(sel + '::'));
      });
    }
    this.s.queue = cards.sort(() => Math.random() - 0.5);
    this.s.qIdx = 0;
    this.s.flipped = false;
    this.s.aiEvalText = '';
    this.s.sessionReviewed = 0;
    this.s.sessionCorrect  = 0;
  },

  // ── Render dispatcher ───────────────────────────────────────────────────────
  _render(view) {
    this.s.view = view;
    const map = {
      loading:  () => this._vLoading(),
      setup:    () => this._vSetup(),
      login:    () => this._vLogin(),
      review:   () => this._vReview(),
      browse:   () => this._vBrowse(),
      decks:    () => this._vDecks(),
      add:      () => this._vAdd(),
      settings: () => this._vSettings(),
      import:   () => this._vImport()
    };
    document.getElementById('app').innerHTML = (map[view] || map.loading)();
    this._bind(view);
  },

  // ── Views ────────────────────────────────────────────────────────────────────
  _vLoading() {
    return `<div class="loading"><div class="spinner"></div><p>Loading…</p></div>`;
  },

  _vSetup() {
    return `
      <div class="login-page">
        <div class="login-card" style="max-width:500px">
          <div class="app-name">PGAnkiADS</div>
          <div class="tagline">First-time setup — enter your credentials below</div>
          ${this._settingsForm(true)}
        </div>
      </div>`;
  },

  _vLogin() {
    return `
      <div class="login-page">
        <div class="login-card">
          <div class="app-name">PGAnkiADS</div>
          <div class="tagline">Anki-powered spaced repetition · Google Sheets + Google Drive + Azure OpenAI</div>
          ${Config.isConfigured()
            ? `<button class="google-btn" id="sign-in-btn">${googleSvg()} Sign in with Google</button>`
            : `<div class="setup-notice">Not configured yet. <a id="goto-setup">Go to Setup →</a></div>`}
        </div>
      </div>`;
  },

  _vReview() {
    const { queue, qIdx, sessionReviewed, selectedTags, tagSearch, tagGroups, answerExpanded, answerVisible,
            reviewLeftTab, selectedDeck, reviewDeckSearch, reviewDeckTreeExpanded, reviewLeftCollapsed } = this.s;
    const total = queue.length;

    // Build tag tree with due counts
    const { nodes: rnodes, roots: rroots } = this._buildTagTree(this.s.cards, this.s.progress);
    const rtf = tagSearch.toLowerCase();
    let reviewTreeMatches = null;
    if (rtf) {
      reviewTreeMatches = new Set();
      for (const [path, node] of Object.entries(rnodes)) {
        if (node.label.toLowerCase().includes(rtf) || node.fullPath.toLowerCase().includes(rtf)) {
          reviewTreeMatches.add(path);
          let p = node.parentPath;
          while (p) { reviewTreeMatches.add(p); p = rnodes[p]?.parentPath; }
        }
      }
    }

    // Build deck tree with due counts
    const { nodes: rdnodes, roots: rdroots } = this._buildDeckTree(this.s.cards, this.s.progress);
    const rdf = reviewDeckSearch.toLowerCase();
    let reviewDeckMatches = null;
    if (rdf) {
      reviewDeckMatches = new Set();
      for (const [path, node] of Object.entries(rdnodes)) {
        if (node.label.toLowerCase().includes(rdf) || node.fullPath.toLowerCase().includes(rdf)) {
          reviewDeckMatches.add(path);
          let p = node.parentPath;
          while (p) { reviewDeckMatches.add(p); p = rdnodes[p]?.parentPath; }
        }
      }
    }

    const hasFilter = selectedTags.length > 0 || selectedDeck.length > 0;
    const tagsBtnLabel = reviewLeftTab === 'decks'
      ? (selectedDeck.length > 0 ? `Decks · ${selectedDeck.length} active ▾` : '⊞ Decks ▾')
      : (selectedTags.length > 0 ? `Tags · ${selectedTags.length} active ▾` : '⊞ Tags ▾');
    const mobileTagsBar = `
      <div class="mobile-bar">
        <button class="mobile-tags-btn${hasFilter ? ' has-active' : ''}" id="tags-drawer-btn">${tagsBtnLabel}</button>
      </div>`;

    const leftCol = `
      <div class="col-tags${this.s.tagsDrawerOpen ? ' drawer-open' : ''}${reviewLeftCollapsed ? ' collapsed' : ''}">

        ${reviewLeftCollapsed ? `
        <button class="panel-strip-btn" id="review-left-toggle" title="Expand panel">▶</button>
        ` : `

        <!-- ── Tab bar: Tags | Decks ── -->
        <div class="left-tab-bar">
          <button class="left-tab${reviewLeftTab !== 'decks' ? ' active' : ''}" data-review-tab="tags">Tags</button>
          <button class="left-tab${reviewLeftTab === 'decks' ? ' active' : ''}" data-review-tab="decks">Decks</button>
          <button class="panel-toggle-btn" id="review-left-toggle" title="Collapse panel">◀</button>
        </div>

        ${reviewLeftTab === 'decks' ? `
        <!-- ── Decks mode ── -->
        <div class="col-tags-top">
          <div class="col-hd"><span>Decks</span></div>
          ${selectedDeck.length > 0 ? `
          <div class="browse-active-tags">
            ${selectedDeck.map(d => `<span class="active-tag-pill">${h(tagDisplayName(d))}<button class="active-deck-rm" data-rm="${h(d)}">✕</button></span>`).join('')}
            <button class="clear-browse-tags" id="clear-decks">Clear all</button>
          </div>` : ''}
          <input class="col-search" id="review-deck-search" placeholder="Search decks…" value="${h(reviewDeckSearch)}">
        </div>
        <div class="col-tags-tree">
          ${rdroots.length === 0
            ? '<div class="col-empty">No decks yet</div>'
            : rdroots
                .filter(r => !reviewDeckMatches || reviewDeckMatches.has(r.fullPath))
                .map(r => this._renderTreeNode(r, 0, selectedDeck, reviewDeckTreeExpanded, reviewDeckMatches, rdnodes,
                    { showDue: true, selAttr: 'dsel', toggleAttr: 'dtoggle', toggleClass: 'deck-toggle' }))
                .join('')}
        </div>
        ` : `
        <!-- ── Tags mode ── -->
        <div class="col-tags-top">
          <div class="col-hd"><span>Tags</span></div>
          ${selectedTags.length > 0 ? `
          <div class="browse-active-tags">
            ${selectedTags.map(t => `<span class="active-tag-pill">${h(tagDisplayName(t))}<button class="active-tag-rm" data-rm="${h(t)}">✕</button></span>`).join('')}
            <button class="clear-browse-tags" id="clear-tags">Clear all</button>
            <button class="tg-save-btn" id="save-tag-group">Save…</button>
          </div>` : ''}
          <input class="col-search" id="tag-search" placeholder="Search tags…" value="${h(tagSearch)}">
        </div>
        <div class="col-tags-tree">
          ${rroots.length === 0
            ? '<div class="col-empty">No tags yet</div>'
            : rroots
                .filter(r => !reviewTreeMatches || reviewTreeMatches.has(r.fullPath))
                .map(r => this._renderTreeNode(r, 0, selectedTags, this.s.reviewTreeExpanded, reviewTreeMatches, rnodes, { showDue: true, showReset: true }))
                .join('')}
        </div>
        <div class="col-tags-groups">
          <div class="col-hd">Groups</div>
          ${tagGroups.filter(g => g.id).map(g => {
            const gTags = parseTags(g.tags);
            let gTotal = 0, gDue = 0;
            for (const c of this.s.cards) {
              const ct = parseTags(c.tags || '');
              if (gTags.some(sel => ct.some(t => t === sel || t.startsWith(sel + '::')))) {
                gTotal++;
                if (SM2.isDue(this.s.progress[c.id])) gDue++;
              }
            }
            return `
              <div class="tg-item">
                <span class="tg-name" title="${h(g.tags)}">${h(g.name)}</span>
                <div class="tg-counts">
                  <span class="tree-cnt">${gTotal}</span>
                  <span class="tree-due${gDue > 0 ? ' has-due' : ''}">${gDue}</span>
                </div>
                <div class="tg-actions">
                  <button class="tg-load-btn" data-gid="${g.id}" title="Filter by group">▶</button>
                  <button class="tg-reset-btn" data-reset-gid="${g.id}" title="Reset group cards">↺</button>
                  <button class="tg-del-btn" data-gid="${g.id}" title="Delete group">✕</button>
                </div>
              </div>`;
          }).join('') || '<div class="col-empty">No groups yet</div>'}
        </div>
        `}
        `}

      </div>`;

    // ── Empty / done states — left col always visible, no right col ──
    if (total === 0) {
      const allCards = this.s.cards.length;
      const tagNote = selectedTags.length > 0
        ? `<p style="color:var(--text2);font-size:13px">Filtered to tags: ${selectedTags.map(t=>`<span class="tag">${h(tagDisplayName(t))}</span>`).join(' ')}</p>`
        : selectedDeck.length > 0
          ? `<p style="color:var(--text2);font-size:13px">Filtered to deck: ${selectedDeck.map(d=>`<span class="tag">${h(tagDisplayName(d))}</span>`).join(' ')}</p>`
          : '';
      return this._withNavFull('review', `
        <div class="review-body">
          ${this.s.tagsDrawerOpen ? '<div class="drawer-backdrop" id="drawer-backdrop"></div>' : ''}
          ${leftCol}
          <div class="col-main">
            ${mobileTagsBar}
            <div class="done-screen">
              <div class="icon">✓</div>
              <h2>All caught up!</h2>
              ${tagNote}
              <p>${allCards ? 'No cards are due for review right now. Come back later or add new cards.' : 'No cards yet — add some to get started.'}</p>
              <div class="done-btns">
                <button class="btn btn-primary" id="go-add">+ Add Cards</button>
                <button class="btn btn-secondary" id="go-browse">Browse Cards</button>
              </div>
            </div>
          </div>
        </div>`);
    }

    if (qIdx >= total) {
      return this._withNavFull('review', `
        <div class="review-body">
          ${this.s.tagsDrawerOpen ? '<div class="drawer-backdrop" id="drawer-backdrop"></div>' : ''}
          ${leftCol}
          <div class="col-main">
            ${mobileTagsBar}
            <div class="done-screen">
              <div class="icon">🎯</div>
              <h2>Session complete!</h2>
              <p>Reviewed ${sessionReviewed} card${sessionReviewed !== 1 ? 's' : ''} · ${this.s.sessionCorrect} correct.</p>
              <div class="done-btns">
                <button class="btn btn-primary" id="restart-btn">Review Again</button>
                <button class="btn btn-secondary" id="go-browse">Browse Cards</button>
              </div>
            </div>
          </div>
        </div>`);
    }

    // ── Active card ──
    const card = queue[qIdx];
    const prog = this.s.progress[card.id] || SM2.defaultProgress();
    const intervals = SM2.previewIntervals(prog);
    const cardTags = parseTags(card.tags || '');
    const pct = Math.round((sessionReviewed / total) * 100);
    const { aiEvalText, aiEvalLoading } = this.s;
    const hasAnswer = !!this.s.userAnswer.trim();
    const showAI = !!(aiEvalText || aiEvalLoading);

    return this._withNavFull('review', `
      <div class="review-body">
        ${this.s.tagsDrawerOpen ? '<div class="drawer-backdrop" id="drawer-backdrop"></div>' : ''}
        ${leftCol}

        <!-- ── Center: Main review column ── -->
        <div class="col-main">
          ${mobileTagsBar}
          <div class="col-main-scroll">
            <div class="review-bar">
              <span class="count">${sessionReviewed}/${total}</span>
              <div class="progress-track"><div class="progress-fill" style="width:${pct}%"></div></div>
              <button class="audio-btn" id="audio-toggle" title="${this.s.audioOn ? 'Mute audio' : 'Enable audio'}">
                ${this.s.audioOn ? '🔊' : '🔇'}
              </button>
            </div>

            <div class="question-card${!answerVisible ? ' clickable' : ''}" id="question-card">
              <div class="side-label">Question</div>
              ${cardTags.length ? `<div class="card-tags">${cardTags.map(t => `<span class="tag">${h(tagDisplayName(t))}</span>`).join('')}</div>` : ''}
              <div class="card-front ${card.front.length > 180 ? 'sm' : ''}">${h(card.front)}</div>
              ${!answerVisible ? `<div class="qa-hint">Click to reveal answer →</div>` : ''}
            </div>

            <div class="answer-area">
              <textarea
                id="user-answer"
                class="answer-input"
                placeholder="Type your answer… or tap 🎤 to speak"
                spellcheck="false"
              >${h(this.s.userAnswer)}</textarea>
              <div class="answer-btns">
                <button class="mic-btn${this.s.isListening ? ' listening' : ''}" id="mic-btn"
                  title="${this.s.isListening ? 'Stop recording' : 'Voice input'}">
                  ${this.s.isListening ? '⏹' : '🎤'}
                </button>
                <button class="ai-validate-btn" id="ai-validate-btn"
                  title="${hasAnswer ? 'Validate my answer with AI' : 'Explain with AI'}"
                  ${aiEvalLoading ? 'disabled' : ''}>
                  ✦
                </button>
              </div>
            </div>

            ${showAI ? `
            <div class="ai-inline-panel">
              <div class="ai-inline-hd">
                <span>✦ ${hasAnswer ? 'AI Evaluation' : 'AI Explanation'}</span>
                <div style="display:flex;gap:4px;align-items:center">
                  ${!aiEvalLoading ? `<button class="ai-eval-close" id="ai-eval-play" title="Read aloud">▶</button>` : ''}
                  <button class="ai-eval-close" id="ai-eval-close" title="Close">✕</button>
                </div>
              </div>
              ${aiEvalLoading
                ? '<div class="ai-eval-loading"><div class="spinner" style="width:24px;height:24px;border-width:2px"></div><span>Thinking…</span></div>'
                : `<div class="ai-inline-body">${h(aiEvalText)}</div>`}
            </div>` : ''}
          </div>

          <div class="rating-bar-col">
            <div class="rating-grid">
              ${[
                {r:1,label:'Again'}, {r:2,label:'Hard'},
                {r:3,label:'Good'},  {r:4,label:'Easy'}
              ].map(({r,label}) => `
                <button class="rating-btn" data-r="${r}">
                  <span>${label}</span>
                  <span class="ivl">${SM2.formatInterval(intervals[r-1])}</span>
                </button>`).join('')}
            </div>
          </div>
        </div>

        <!-- ── Right: Answer column — only rendered when revealed ── -->
        ${answerVisible ? `
        <div class="col-answer${answerExpanded ? ' expanded' : ''}">
          <div class="col-hd">
            <span>Answer <span class="edit-hint">· dbl-click to edit</span></span>
            <div style="display:flex;gap:4px;align-items:center">
              <button class="col-expand-btn" id="answer-play" title="Read aloud">▶</button>
              <button class="col-expand-btn" id="answer-col-close" title="Close answer">✕</button>
              <button class="col-expand-btn" id="expand-answer" title="${answerExpanded ? 'Collapse' : 'Expand'}">
                ${answerExpanded ? '⊠' : '⤢'}
              </button>
            </div>
          </div>

          <div class="answer-editor-toolbar" id="editor-toolbar">
            <button class="tb-btn" data-cmd="bold"                 title="Bold"><b>B</b></button>
            <button class="tb-btn" data-cmd="italic"               title="Italic"><i>I</i></button>
            <button class="tb-btn" data-cmd="underline"            title="Underline"><u>U</u></button>
            <button class="tb-btn" data-cmd="strikeThrough"        title="Strikethrough"><s>S</s></button>
            <div class="tb-sep"></div>
            <button class="tb-btn" data-cmd="insertUnorderedList"  title="Bullet list">• List</button>
            <button class="tb-btn" data-cmd="insertOrderedList"    title="Numbered">1. List</button>
            <button class="tb-btn" data-cmd="insertHorizontalRule" title="Divider">—</button>
            <div class="tb-sep"></div>
            <button class="tb-btn" id="tb-code"  title="Code">&lt;/&gt;</button>
            <div class="tb-sep"></div>
            <button class="tb-btn" id="tb-html"  title="Toggle raw HTML source">HTML</button>
            <div style="flex:1"></div>
            <button class="tb-btn tb-cancel" id="tb-cancel">Cancel</button>
            <button class="tb-btn tb-save"   id="tb-save">Save</button>
          </div>

          <div class="correct-answer-html" id="answer-html"></div>
        </div>` : ''}

      </div>`);
  },

  _vBrowse() {
    const { browseFilterTags, browseSelectedCard, browseTagSearch, browseSearch, tagTreeExpanded,
            browseDeckFilter, browseLeftTab, browseDeckSearch, deckTreeExpanded, browseSelectedCards,
            browseLeftCollapsed } = this.s;

    // Build tag tree
    const { nodes, roots } = this._buildTagTree(this.s.cards);
    const lf = (browseTagSearch || '').toLowerCase();
    let treeMatches = null;
    if (lf) {
      treeMatches = new Set();
      for (const [path, node] of Object.entries(nodes)) {
        if (node.label.toLowerCase().includes(lf) || node.fullPath.toLowerCase().includes(lf)) {
          treeMatches.add(path);
          let p = node.parentPath;
          while (p) { treeMatches.add(p); p = nodes[p]?.parentPath; }
        }
      }
    }

    // Build deck tree
    const { nodes: dnodes, roots: droots } = this._buildDeckTree(this.s.cards);
    const df = (browseDeckSearch || '').toLowerCase();
    let deckTreeMatches = null;
    if (df) {
      deckTreeMatches = new Set();
      for (const [path, node] of Object.entries(dnodes)) {
        if (node.label.toLowerCase().includes(df) || node.fullPath.toLowerCase().includes(df)) {
          deckTreeMatches.add(path);
          let p = node.parentPath;
          while (p) { deckTreeMatches.add(p); p = dnodes[p]?.parentPath; }
        }
      }
    }

    // Filter cards — tags OR deck (whichever is active)
    let filtered = this.s.cards;
    if (browseFilterTags.length > 0) {
      filtered = filtered.filter(card => {
        const ct = parseTags(card.tags || '');
        return browseFilterTags.some(sel => ct.some(t => t === sel || t.startsWith(sel + '::')));
      });
    }
    if (browseDeckFilter.length > 0) {
      filtered = filtered.filter(card => {
        const d = (card.deck || 'Default').trim();
        return browseDeckFilter.some(sel => d === sel || d.startsWith(sel + '::'));
      });
    }
    const bq = (browseSearch || '').toLowerCase();
    if (bq) filtered = filtered.filter(c => `${c.front} ${c.back} ${c.tags} ${c.deck||''}`.toLowerCase().includes(bq));

    const hasFilter = browseFilterTags.length > 0 || browseDeckFilter.length > 0;
    const browseFilterBtnLabel = hasFilter
      ? `Filter · active ▾`
      : '⊞ Filter ▾';
    const selSet = new Set(browseSelectedCards);

    // All existing deck names for datalist in assign prompt alternative
    const allDecks = [...new Set(this.s.cards.map(c => c.deck || 'Default'))].sort();

    return this._withNavFull('browse', `
      <div class="browse-body">
        ${this.s.browseTagsDrawerOpen ? '<div class="drawer-backdrop" id="browse-backdrop"></div>' : ''}

        <!-- ── Left: Tags / Decks ── -->
        <div class="browse-col-tags${this.s.browseTagsDrawerOpen ? ' drawer-open' : ''}${browseLeftCollapsed ? ' collapsed' : ''}">
          ${browseLeftCollapsed ? `
          <button class="panel-strip-btn" id="browse-left-toggle" title="Expand panel">▶</button>
          ` : `
          <div class="left-tab-bar">
            <button class="left-tab${browseLeftTab !== 'decks' ? ' active' : ''}" data-browse-tab="tags">Tags</button>
            <button class="left-tab${browseLeftTab === 'decks' ? ' active' : ''}" data-browse-tab="decks">Decks</button>
            <button class="panel-toggle-btn" id="browse-left-toggle" title="Collapse panel">◀</button>
          </div>
          <div class="browse-left-content">
            ${browseLeftTab === 'decks' ? `
            ${browseDeckFilter.length > 0 ? `
            <div class="browse-active-tags">
              ${browseDeckFilter.map(d => `
                <span class="active-tag-pill">${h(tagDisplayName(d))}<button class="active-deck-rm" data-rm="${h(d)}">✕</button></span>
              `).join('')}
              <button class="clear-browse-tags" id="clear-browse-decks">Clear all</button>
            </div>` : ''}
            <input class="col-search" id="browse-deck-search" placeholder="Search decks…" value="${h(browseDeckSearch||'')}">
            <div class="tag-tree">
              ${droots.length === 0
                ? '<div class="col-empty">No decks yet</div>'
                : droots
                    .filter(r => !deckTreeMatches || deckTreeMatches.has(r.fullPath))
                    .map(r => this._renderTreeNode(r, 0, browseDeckFilter, deckTreeExpanded, deckTreeMatches, dnodes,
                        { selAttr: 'dsel', toggleAttr: 'dtoggle', toggleClass: 'deck-toggle' }))
                    .join('')}
            </div>
            ` : `
            ${browseFilterTags.length > 0 ? `
            <div class="browse-active-tags">
              ${browseFilterTags.map(t => `
                <span class="active-tag-pill">${h(tagDisplayName(t))}<button class="active-tag-rm" data-rm="${h(t)}">✕</button></span>
              `).join('')}
              <button class="clear-browse-tags" id="clear-browse-tags">Clear all</button>
            </div>` : ''}
            <input class="col-search" id="browse-tag-search" placeholder="Search tags…" value="${h(browseTagSearch||'')}">
            <div class="tag-tree">
              ${roots.length === 0
                ? '<div class="col-empty">No tags yet</div>'
                : roots
                    .filter(r => !treeMatches || treeMatches.has(r.fullPath))
                    .map(r => this._renderTreeNode(r, 0, browseFilterTags, tagTreeExpanded, treeMatches, nodes))
                    .join('')}
            </div>
            `}
          </div>
          `}
        </div>

        <!-- ── Center: Card list ── -->
        <div class="browse-col-cards">
          <div class="mobile-browse-bar">
            <button class="mobile-filter-btn${hasFilter ? ' has-active' : ''}" id="browse-filter-btn">${browseFilterBtnLabel}</button>
          </div>
          ${browseSelectedCards.length > 0 ? `
          <div class="bulk-bar">
            <span class="bulk-count">${browseSelectedCards.length} selected</span>
            <button class="btn btn-secondary btn-sm" id="bulk-assign-deck">Move to Deck…</button>
            <button class="btn btn-sm" id="bulk-clear-sel" style="background:transparent;border:none;color:var(--text2);cursor:pointer">✕ Clear</button>
          </div>` : ''}
          <div class="browse-cards-hd">
            <span>Cards <span class="count-badge">(${filtered.length})</span></span>
            <div class="browse-hd-right">
              <input class="col-search" id="browse-search" placeholder="Search cards…" value="${h(browseSearch||'')}">
              <button class="btn btn-primary btn-sm" id="go-add">+ Add</button>
              ${Config.azureApiKey ? `<button class="btn btn-secondary btn-sm" id="ai-gen-btn">✦ AI</button>` : ''}
            </div>
          </div>
          <div class="browse-card-list">
            ${filtered.length === 0
              ? `<div class="col-empty" style="padding:32px;text-align:center">No cards found</div>`
              : filtered.map(card => {
                  const prog = this.s.progress[card.id];
                  const due = SM2.isDue(prog);
                  const sel = browseSelectedCard?.id === card.id;
                  const checked = selSet.has(card.id);
                  const ct = parseTags(card.tags || '');
                  const deck = card.deck || 'Default';
                  return `
                    <div class="bci${sel ? ' sel' : ''}${checked ? ' checked' : ''}">
                      <label class="bci-check" title="Select card">
                        <input type="checkbox" class="bci-cb" data-cbid="${card.id}"${checked ? ' checked' : ''}>
                      </label>
                      <div class="bci-main" data-cid="${card.id}">
                        <div class="bci-q">${h(card.front.slice(0,120))}${card.front.length>120?'…':''}</div>
                        <div class="bci-deck-row"><span class="deck-badge" title="${h(deck)}">◈ ${h(tagDisplayName(deck))}</span></div>
                        <div class="bci-foot">
                          <div class="bci-tags">${ct.map(t=>`<span class="tag">${h(tagDisplayName(t))}</span>`).join('')}</div>
                          <div class="bci-acts">
                            <span class="bci-due" style="color:${due?'var(--danger)':'var(--text2)'}">${due?'Due':(prog?.dueDate||'New')}</span>
                            <button class="icon-btn" data-edit="${card.id}" title="Edit">✎</button>
                            <button class="icon-btn del" data-del="${card.id}" title="Delete">✕</button>
                          </div>
                        </div>
                      </div>
                    </div>`;
                }).join('')}
          </div>
        </div>

        <!-- ── Right: Card detail — only visible when a card is selected ── -->
        ${browseSelectedCard ? `
        <div class="browse-col-detail has-selection">
          <div class="browse-detail-q">
            <div class="col-hd">
              <span>Question</span>
              <button class="detail-close-btn" id="browse-detail-close">✕</button>
            </div>
            <div class="browse-detail-text">${h(browseSelectedCard.front)}</div>
          </div>
          <div class="browse-detail-a">
            <div class="col-hd">Answer</div>
            <div class="browse-answer-html" id="browse-answer-html"></div>
          </div>
        </div>` : ''}

      </div>`);
  },

  _vDecks() {
    const { deckManagerExpanded, decksShowNewForm, decksNewInput } = this.s;
    const { nodes, roots } = this._buildDeckTree(this.s.cards);
    const totalDecks = Object.keys(nodes).length;

    return this._withNav('decks', `
      <div class="decks-page">
        <div class="decks-pg-header">
          <h2>Decks</h2>
          ${totalDecks > 0 ? `<span class="count-badge">(${totalDecks} ${totalDecks === 1 ? 'deck' : 'decks'})</span>` : ''}
          <button class="btn btn-primary btn-sm" id="deck-new-btn">+ New Deck</button>
        </div>

        ${decksShowNewForm ? `
        <div class="deck-new-form">
          <input class="deck-new-input" id="deck-new-input" type="text"
            placeholder="Deck name — use :: for hierarchy (e.g. Java::Collections)"
            value="${h(decksNewInput)}" autocomplete="off">
          <div class="deck-new-actions">
            <button class="btn btn-primary btn-sm" id="deck-new-upload-btn" title="Upload CSV/TSV cards to this deck">↑ Upload Cards</button>
            <input type="file" accept=".csv,.tsv,.txt" id="deck-new-file">
            <button class="btn btn-secondary btn-sm" id="deck-new-add-manual">+ Add Card</button>
            <button class="btn btn-sm deck-cancel-btn" id="deck-create-cancel">Cancel</button>
          </div>
        </div>` : ''}

        <div class="deck-tree-mgr">
          ${roots.length === 0
            ? `<div class="deck-tree-mgr-empty">
                 <div style="font-size:32px;margin-bottom:8px">⊟</div>
                 <div style="font-weight:600;margin-bottom:4px">No decks yet</div>
                 <div style="color:var(--text2);font-size:13px">Click "+ New Deck" to create one and upload cards.</div>
               </div>`
            : roots.map(r => this._renderDeckManagerNode(r, 0, deckManagerExpanded, nodes)).join('')}
        </div>

        <div class="deck-upload-hint">
          <span>Upload: CSV (comma), TSV (tab), or space-separated with columns</span>
          <code>front</code><code>back</code><code>tags</code><code>notes</code>
          <span>— first row must be headers. Delimiter auto-detected.</span>
        </div>
      </div>`);
  },

  _renderDeckManagerNode(node, depth, expanded, nodes) {
    const children = [...node.childPaths]
      .map(p => nodes[p]).filter(Boolean)
      .sort((a, b) => a.label.localeCompare(b.label));
    const hasKids = children.length > 0;
    const isExpanded = !expanded[node.fullPath]; // default: expanded

    return `
      <div class="deck-mgr-node">
        <div class="deck-mgr-row" style="padding-left:${depth * 20 + 8}px">
          ${hasKids
            ? `<button class="deck-mgr-toggle" data-dmtoggle="${h(node.fullPath)}" title="${isExpanded ? 'Collapse' : 'Expand'}">${isExpanded ? '−' : '+'}</button>`
            : `<span class="deck-indent"></span>`}
          <button class="deck-mgr-name" data-dmnav="${h(node.fullPath)}">${h(node.label)}</button>
          <span class="deck-count-badge" title="${node.cardIds.size} card(s)">${node.cardIds.size}</span>
          <div class="deck-mgr-actions">
            <button class="deck-upload-lbl" data-uploadbtn="${h(node.fullPath)}" title="Upload CSV/TSV cards to ${h(tagDisplayName(node.fullPath))}">↑ Upload</button>
            <input type="file" accept=".csv,.tsv,.txt" class="deck-file-input" data-deck="${h(node.fullPath)}">
            <button class="deck-rename-btn" data-dmrename="${h(node.fullPath)}" title="Rename deck">✎</button>
          </div>
        </div>
        ${hasKids && isExpanded ? `
          <div class="deck-mgr-kids">
            ${children.map(c => this._renderDeckManagerNode(c, depth + 1, expanded, nodes)).join('')}
          </div>` : ''}
      </div>`;
  },

  _vAdd() {
    const card = this.s.editCard || {};
    const isEdit = !!card.id;
    return this._withNav('browse', `
      <div class="panel" style="max-width:640px">
        <h2>${isEdit ? 'Edit Card' : 'Add Card'}</h2>
        <br>
        <div class="form-group">
          <label>Question (Front)</label>
          <textarea id="f-front">${h(card.front || '')}</textarea>
        </div>
        <div class="form-group">
          <label>Answer (Back) <span style="font-weight:400;text-transform:none">(optional)</span></label>
          <textarea id="f-back" placeholder="Leave blank to add the answer later">${h(card.back || '')}</textarea>
        </div>
        <div class="form-group">
          <label>Deck</label>
          <input id="f-deck" type="text" list="f-deck-opts" placeholder="Default" value="${h(card.deck || 'Default')}">
          <datalist id="f-deck-opts">
            ${[...new Set(this.s.cards.map(c => c.deck || 'Default'))].sort().map(d => `<option value="${h(d)}">`).join('')}
          </datalist>
          <div class="hint">One deck per card. Use <code>::</code> for hierarchy (e.g. <code>Java::Collections</code>). Defaults to <code>Default</code>.</div>
        </div>
        <div class="form-group">
          <label>Tags</label>
          <input id="f-tags" type="text" placeholder="java::collections :: Spring Boot :: algorithms" value="${h(card.tags || '')}">
          <div class="hint">Separate tags with <code> :: </code> · Use <code>::</code> within a tag for hierarchy (e.g. <code>java::collections::List</code>)</div>
        </div>
        <div class="form-group">
          <label>Notes (optional)</label>
          <textarea id="f-notes" style="min-height:60px">${h(card.notes || '')}</textarea>
        </div>
        <div class="form-actions">
          <button class="btn btn-primary" id="save-btn" data-id="${card.id||''}">${isEdit ? 'Save Changes' : 'Add Card'}</button>
          <button class="btn btn-secondary" id="cancel-btn">Cancel</button>
        </div>
      </div>`);
  },

  _vSettings() {
    return this._withNav('settings', `
      <div class="panel" style="max-width:600px">
        <h2>Settings</h2>
        <p class="sub">Configure Google API credentials and Azure OpenAI connection.</p>
        ${this._settingsForm(false)}
      </div>`);
  },

  // ── Shared templates ─────────────────────────────────────────────────────────
  _settingsForm(isSetup) {
    const wrap = isSetup ? 'style="border:none;background:transparent;padding:0"' : '';
    return `
      <div class="panel" ${wrap}>
        <div class="form-group">
          <label>Google OAuth Client ID</label>
          <input id="s-gci" type="text" placeholder="123456.apps.googleusercontent.com" value="${h(Config.googleClientId)}">
          <div class="hint">Google Cloud Console → APIs &amp; Services → Credentials → OAuth 2.0 Client ID</div>
        </div>
        <div class="form-group">
          <label>Google API Key <span style="font-weight:400;text-transform:none">(optional)</span></label>
          <input id="s-gak" type="password" placeholder="AIza… (leave blank — OAuth token is sufficient)" value="${h(Config.googleApiKey)}">
          <div class="hint">Only needed if you want public/unauthenticated read access. Leave blank to use OAuth only.</div>
        </div>
        <div class="form-group">
          <label>Google Sheet ID <span style="color:var(--danger);font-weight:400">*</span></label>
          ${Config.sheetId
            ? `<div class="sheet-status ok">✓ Connected — <a class="sheet-link" href="https://docs.google.com/spreadsheets/d/${h(Config.sheetId)}/edit" target="_blank" rel="noopener">open sheet ↗</a></div>`
            : `<div class="sheet-status warn">⚠ No sheet connected — paste your Sheet ID below</div>`}
          <input id="s-sid" type="text" placeholder="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms"
                 value="${h(Config.sheetId)}"
                 ${window.ENV_SHEET_ID ? 'readonly title="Injected from SHEET_ID repository secret"' : ''}>
          <div class="hint">
            ${window.ENV_SHEET_ID
              ? 'Set via the <code>SHEET_ID</code> repository secret — change it in GitHub Settings → Secrets &amp; variables → Actions.'
              : 'Copy from your Google Sheet URL: docs.google.com/spreadsheets/d/<b>SHEET_ID_HERE</b>/edit'}
          </div>
        </div>
        <div class="form-group">
          <label>Theme</label>
          <div class="theme-picker">
            ${THEMES.map(t => `
              <button class="theme-swatch${Config.theme===t.id?' sel':''}" data-theme-pick="${t.id}" type="button">
                <span class="ts-preview" style="background:${t.bg};border-color:${t.accent}">
                  <span class="ts-dot" style="background:${t.accent}"></span>
                  <span class="ts-dot" style="background:${t.bg}; border:1px solid ${t.accent}"></span>
                </span>
                <span class="ts-label">${t.label}</span>
              </button>`).join('')}
          </div>
        </div>
        <div class="form-group">
          <label>Azure TTS Endpoint <span style="font-weight:400;text-transform:none">(if different from chat endpoint)</span></label>
          <input id="s-aztep" type="text" placeholder="https://….cognitiveservices.azure.com/" value="${h(Config.azureTtsEndpoint)}">
          <div class="hint">Leave blank to use the same endpoint as chat.</div>
        </div>
        <div class="form-group">
          <label>Azure TTS API Key <span style="font-weight:400;text-transform:none">(if different from chat key)</span></label>
          <input id="s-aztak" type="password" placeholder="TTS resource api-key…" value="${h(Config.azureTtsApiKey)}">
          <div class="hint">Leave blank to use the same key as chat. Can also be set in the Sheet's Settings tab as <code>azureTtsApiKey</code>.</div>
        </div>
        <div class="form-group">
          <label>Azure TTS Deployment <span style="font-weight:400;text-transform:none">(optional — for natural AI voices)</span></label>
          <input id="s-aztts" type="text" placeholder="gpt-4o-mini-tts" value="${h(Config.azureTtsDeployment)}">
          <div class="hint">Deployment name for the TTS model (e.g. <code>gpt-4o-mini-tts</code>). Uses <code>Authorization: Bearer</code> auth.</div>
        </div>
        <div class="form-group">
          <label>TTS API Version <span style="font-weight:400;text-transform:none">(optional)</span></label>
          <input id="s-azttsv" type="text" placeholder="2025-03-01-preview" value="${h(Config.azureTtsApiVersion)}">
          <div class="hint">Default: <code>2025-03-01-preview</code>.</div>
        </div>
        <div class="form-group">
          <label>Voice</label>
          <div class="tts-voice-grid">
            ${TTS_VOICES.map(v => `
              <button class="tts-voice-btn${Config.ttsVoice===v.id?' sel':''}" data-ttsv="${v.id}" type="button">
                <span class="ttsv-label">${v.label}</span>
                <span class="ttsv-desc">${v.desc}</span>
              </button>`).join('')}
          </div>
          ${Config.azureTtsDeployment
            ? `<div class="hint">Using Azure AI voice — click to preview.</div>`
            : `<div class="hint">Set an Azure TTS Deployment above to use AI voices. Using browser fallback until then.</div>`}
          <details style="margin-top:10px">
            <summary style="font-size:12px;color:var(--text2);cursor:pointer">Browser fallback voice</summary>
            <div class="voice-list" id="voice-list" style="margin-top:8px">
              <div class="voice-loading">Loading voices…</div>
            </div>
          </details>
        </div>
        ${Auth.isSignedIn() ? `
        <div class="sheet-settings-info">
          <div class="sheet-settings-info-hd">
            ✦ Azure settings live in your Sheet's <b>Settings</b> tab
            ${Config.sheetId
              ? `— <a href="https://docs.google.com/spreadsheets/d/${h(Config.sheetId)}/edit#gid=0" target="_blank" rel="noopener">open sheet ↗</a>`
              : ''}
          </div>
          <div class="sheet-settings-info-body">
            Add or edit these rows in the <b>Settings</b> tab (columns: <code>key</code> | <code>value</code>):
            <code class="key-list">azureApiKey · azureEndpoint · azureDeployment · azureApiVersion</code>
            <code class="key-list">azureTtsApiKey · azureTtsEndpoint · azureTtsDeployment · azureTtsApiVersion</code>
            Values are loaded on sign-in and can be reloaded with the button below.
          </div>
          <button class="btn btn-secondary btn-sm" id="reload-from-sheet" style="margin-top:8px">↺ Reload from Sheet</button>
        </div>` : ''}
        <div class="form-group">
          <label>Azure OpenAI API Key</label>
          <input id="s-azk" type="password" placeholder="Azure api-key…" value="${h(Config.azureApiKey)}">
          <div class="hint">Stored in your Sheet's Settings tab as <code>azureApiKey</code> and cached in localStorage.</div>
        </div>
        <div class="form-group">
          <label>Azure Endpoint</label>
          <input id="s-aze" type="text" placeholder="https://…openai.azure.com/" value="${h(Config.azureEndpoint)}">
        </div>
        <div class="form-group">
          <label>Deployment</label>
          <input id="s-azd" type="text" placeholder="gpt-5.1" value="${h(Config.azureDeployment)}">
        </div>
        <div class="form-group">
          <label>API Version</label>
          <input id="s-azv" type="text" placeholder="2024-12-01-preview" value="${h(Config.azureApiVersion)}">
        </div>
        <div class="form-actions">
          <button class="btn btn-primary" id="save-settings">Save Settings</button>
          ${!isSetup && Auth.isSignedIn()
            ? `<button class="btn btn-danger" id="sign-out-btn">Sign Out</button>`
            : ''}
        </div>
      </div>`;
  },

  _themeDots() {
    const cur = Config.theme;
    return `<div class="theme-dots">
      ${THEMES.map(t => `
        <button class="theme-dot${cur===t.id?' active':''}" data-theme-dot="${t.id}"
          title="${t.label}" style="background:${t.dot}"></button>`).join('')}
    </div>`;
  },

  _withNav(active, body) {
    const user = Auth.user;
    const initial = (user?.name || user?.email || '?')[0].toUpperCase();
    return `
      <div class="layout">
        <div class="topbar">
          <div class="logo">PGAnkiADS</div>
          <div class="topbar-desktop">
            <nav>
              ${['review','browse','decks','import','settings'].map(v => `
                <button class="nav-btn ${active===v?'active':''}" data-nav="${v}">
                  <span>${v[0].toUpperCase()+v.slice(1)}</span>
                </button>`).join('')}
            </nav>
            ${this._themeDots()}
          </div>
          ${user ? `<div class="avatar-chip" title="${user.email||''}">${initial}</div>` : ''}
        </div>
        <main class="main">${body}</main>
        ${this._bottomNav(active)}
      </div>`;
  },

  // Full-width variant — no .main wrapper, used for 3-column review layout
  _withNavFull(active, body) {
    const user = Auth.user;
    const initial = (user?.name || user?.email || '?')[0].toUpperCase();
    return `
      <div class="layout">
        <div class="topbar">
          <div class="logo">PGAnkiADS</div>
          <div class="topbar-desktop">
            <nav>
              ${['review','browse','decks','import','settings'].map(v => `
                <button class="nav-btn ${active===v?'active':''}" data-nav="${v}">
                  <span>${v[0].toUpperCase()+v.slice(1)}</span>
                </button>`).join('')}
            </nav>
            ${this._themeDots()}
          </div>
          ${user ? `<div class="avatar-chip" title="${user.email||''}">${initial}</div>` : ''}
        </div>
        ${body}
        ${this._bottomNav(active)}
      </div>`;
  },

  _bottomNav(active) {
    const items = [
      { id: 'review',   label: 'Review',   icon: '◎' },
      { id: 'browse',   label: 'Browse',   icon: '≡' },
      { id: 'decks',    label: 'Decks',    icon: '⊟' },
      { id: 'import',   label: 'Import',   icon: '⇪' },
      { id: 'settings', label: 'Settings', icon: '⚙' }
    ];
    return `
      <nav class="bottom-nav">
        ${items.map(i => `
          <button class="bn-btn${active===i.id?' active':''}" data-nav="${i.id}">
            <span class="bn-icon">${i.icon}</span>
            <span class="bn-label">${i.label}</span>
          </button>`).join('')}
      </nav>`;
  },

  // ── Tag tree helpers ─────────────────────────────────────────────────────────

  // Build a trie from all card tags. Each tag path (java::collections::list) adds
  // nodes for every prefix. cardIds at each node = all cards with that prefix.
  // Pass progress to also get dueIds (cards currently due) per node.
  _buildTagTree(cards, progress = null) {
    const nodes = {};
    const ensure = (path, parentPath) => {
      if (!nodes[path]) nodes[path] = {
        label: tagSegments(path).at(-1),
        fullPath: path,
        parentPath: parentPath || null,
        cardIds: new Set(),
        dueIds: new Set(),
        childPaths: new Set()
      };
    };
    for (const card of cards) {
      const isDue = progress ? SM2.isDue(progress[card.id]) : false;
      for (const tagPath of parseTags(card.tags || '')) {
        const segs = tagSegments(tagPath);
        let cur = '';
        for (const seg of segs) {
          const parent = cur;
          cur = cur ? `${cur}::${seg}` : seg;
          ensure(cur, parent || null);
          nodes[cur].cardIds.add(card.id);
          if (isDue) nodes[cur].dueIds.add(card.id);
          if (parent) nodes[parent].childPaths.add(cur);
        }
      }
    }
    const roots = Object.values(nodes).filter(n => !n.parentPath)
      .sort((a, b) => a.label.localeCompare(b.label));
    return { nodes, roots };
  },

  // Build deck hierarchy tree. Each card belongs to exactly one deck (card.deck || 'Default').
  _buildDeckTree(cards, progress = null) {
    const nodes = {};
    const ensure = (path, parentPath) => {
      if (!nodes[path]) nodes[path] = {
        label: tagSegments(path).at(-1),
        fullPath: path,
        parentPath: parentPath || null,
        cardIds: new Set(),
        dueIds: new Set(),
        childPaths: new Set()
      };
    };
    for (const card of cards) {
      const isDue = progress ? SM2.isDue(progress[card.id]) : false;
      const deck = (card.deck || 'Default').trim();
      const segs = tagSegments(deck);
      let cur = '';
      for (const seg of segs) {
        const parent = cur;
        cur = cur ? `${cur}::${seg}` : seg;
        ensure(cur, parent || null);
        nodes[cur].cardIds.add(card.id);
        if (isDue) nodes[cur].dueIds.add(card.id);
        if (parent) nodes[parent].childPaths.add(cur);
      }
    }
    const roots = Object.values(nodes).filter(n => !n.parentPath)
      .sort((a, b) => a.label.localeCompare(b.label));
    return { nodes, roots };
  },

  // opts: { showDue, showReset, selAttr, toggleAttr, toggleClass }
  // selAttr/toggleAttr/toggleClass default to tag-tree values; override for deck tree
  _renderTreeNode(node, depth, selectedPaths, expanded, matches, nodes, opts = {}) {
    const {
      showDue = false, showReset = false,
      selAttr = 'tsel', toggleAttr = 'toggle', toggleClass = 'tree-toggle'
    } = opts;
    const children = [...node.childPaths]
      .filter(p => !matches || matches.has(p))
      .map(p => nodes[p])
      .sort((a, b) => a.label.localeCompare(b.label));
    const hasKids = children.length > 0;
    const isExpanded = matches ? true : !expanded[node.fullPath];
    const isSel = selectedPaths.includes(node.fullPath);
    const dueCount = node.dueIds?.size || 0;

    return `
      <div class="tree-node-wrap">
        <div class="tree-row" style="padding-left:${depth * 16}px">
          ${hasKids
            ? `<button class="${toggleClass}" data-${toggleAttr}="${h(node.fullPath)}">${isExpanded ? '▾' : '▸'}</button>`
            : `<span class="tree-indent"></span>`}
          <button class="tree-lbl${isSel ? ' active' : ''}" data-${selAttr}="${h(node.fullPath)}">
            <span class="tree-tag">${h(node.label)}</span>
            <span class="tree-cnt">${node.cardIds.size}</span>
            ${showDue ? `<span class="tree-due${dueCount > 0 ? ' has-due' : ''}">${dueCount}</span>` : ''}
          </button>
          ${showReset ? `<button class="tree-reset-btn" data-reset-tag="${h(node.fullPath)}" title="Reset '${h(node.label)}' cards">↺</button>` : ''}
        </div>
        ${hasKids && isExpanded ? `
          <div class="tree-kids">
            ${children.map(c => this._renderTreeNode(c, depth + 1, selectedPaths, expanded, matches, nodes, opts)).join('')}
          </div>` : ''}
      </div>`;
  },

  // ── Import view ──────────────────────────────────────────────────────────────
  _vImport() {
    return this._withNav('import', `
      <div class="panel" style="max-width:760px;margin:0 auto">
        <h2>Import Anki Cards</h2>
        <p class="sub">Upload a <code>.colpkg</code> or <code>.apkg</code> file exported from Anki.
          Cards will be written to the Google Sheet; media (images, audio, HTML) will be
          uploaded to Google Drive and linked automatically.</p>

        <div class="form-group">
          <label>Anki Package File (.colpkg / .apkg)</label>
          <div id="import-drop-zone" style="
            border:2px dashed var(--border);border-radius:10px;padding:40px 20px;
            text-align:center;cursor:pointer;transition:border-color .15s;color:var(--text2)">
            <div style="font-size:32px;margin-bottom:10px">⇪</div>
            <div>Drag &amp; drop or <a id="import-file-link" style="color:var(--primary);cursor:pointer">click to browse</a></div>
            <div style="font-size:12px;margin-top:6px" id="import-file-name">No file selected</div>
            <input type="file" id="import-file" accept=".colpkg,.apkg" style="display:none">
          </div>
        </div>

        <div class="form-group">
          <label>Deck filter (optional)</label>
          <input type="text" id="import-deck-filter" class="form-control"
            placeholder="e.g. Japanese or leave blank to import all decks"
            style="padding:9px 13px;background:var(--bg3);border:1px solid var(--border);border-radius:8px;color:var(--text);font-size:14px;width:100%">
          <p class="hint">Only cards whose deck name <em>starts with</em> this string will be imported.</p>
        </div>

        <div class="form-group">
          <label>Media handling</label>
          <select id="import-media-mode" style="width:100%;padding:9px 13px;background:var(--bg3);border:1px solid var(--border);border-radius:8px;color:var(--text);font-size:14px">
            <option value="drive">Upload media to Google Drive (recommended)</option>
            <option value="skip">Skip media — import text only</option>
          </select>
        </div>

        <div class="form-actions" style="margin-top:8px">
          <button class="btn btn-primary" id="import-start-btn" disabled>Import</button>
          <button class="btn btn-secondary" id="import-cancel-btn">Cancel</button>
        </div>

        <div id="import-progress" style="display:none;margin-top:24px">
          <div style="display:flex;align-items:center;gap:12px;margin-bottom:12px">
            <div class="spinner" style="width:22px;height:22px;border-width:2px"></div>
            <span id="import-status-msg" style="font-size:14px;color:var(--text2)">Preparing…</span>
          </div>
          <div style="background:var(--bg3);border-radius:6px;height:8px;overflow:hidden">
            <div id="import-progress-bar" style="height:100%;background:var(--primary);width:0%;transition:width .3s"></div>
          </div>
          <pre id="import-log" style="margin-top:12px;max-height:240px;overflow-y:auto;
            background:var(--bg);border:1px solid var(--border);border-radius:8px;
            padding:10px 12px;font-size:12px;color:var(--text2);white-space:pre-wrap"></pre>
        </div>

        <div id="import-result" style="display:none;margin-top:24px">
          <div id="import-result-msg" style="font-size:15px;font-weight:600;margin-bottom:10px"></div>
          <button class="btn btn-primary" id="import-go-browse">Browse Imported Cards</button>
          <button class="btn btn-secondary" id="import-again-btn" style="margin-left:10px">Import Another</button>
        </div>
      </div>`);
  },

  // ── Event binding ────────────────────────────────────────────────────────────
  _bind(view) {
    // Nav links (present on all views with topbar)
    document.querySelectorAll('[data-nav]').forEach(b =>
      b.addEventListener('click', () => {
        if (b.dataset.nav === 'review') this.s.answerVisible = false;
        this._render(b.dataset.nav);
        if (b.dataset.nav === 'review') this._speakCurrent();
      }));

    // Theme dots in topbar — present on all views
    document.querySelectorAll('[data-theme-dot]').forEach(b =>
      b.addEventListener('click', () => {
        this._applyTheme(b.dataset.themeDot, true);
        // Re-render topbar dots without a full view re-render
        document.querySelectorAll('[data-theme-dot]').forEach(d => {
          d.classList.toggle('active', d.dataset.themeDot === b.dataset.themeDot);
        });
      }));

    // Theme swatches in settings form
    document.querySelectorAll('[data-theme-pick]').forEach(b =>
      b.addEventListener('click', () => {
        this._applyTheme(b.dataset.themePick, true);
        document.querySelectorAll('[data-theme-pick]').forEach(s => {
          s.classList.toggle('sel', s.dataset.themePick === b.dataset.themePick);
        });
        document.querySelectorAll('[data-theme-dot]').forEach(d => {
          d.classList.toggle('active', d.dataset.themeDot === b.dataset.themePick);
        });
      }));

    if (view === 'login') {
      document.getElementById('sign-in-btn')?.addEventListener('click', () => this._signIn());
      document.getElementById('goto-setup')?.addEventListener('click', () => this._render('setup'));
    }

    if (view === 'setup' || view === 'settings') {
      document.getElementById('save-settings')?.addEventListener('click', () => this._saveSettings());
      document.getElementById('sign-out-btn')?.addEventListener('click', () => this._signOut());
      document.getElementById('reload-from-sheet')?.addEventListener('click', () => this._reloadFromSheet());
      this._populateVoices();
      document.querySelectorAll('[data-ttsv]').forEach(btn => {
        btn.addEventListener('click', () => {
          const voice = btn.dataset.ttsv;
          Config.ttsVoice = voice;
          // Persist immediately to sheet — don't wait for Save Settings
          if (Auth.isSignedIn()) {
            Sheets.saveSetting('ttsVoice', voice).catch(() => {});
          }
          document.querySelectorAll('[data-ttsv]').forEach(b => b.classList.toggle('sel', b.dataset.ttsv === voice));
          // Read all live inputs — user may not have saved yet
          const formDep = document.getElementById('s-aztts')?.value.trim();
          const formEp  = document.getElementById('s-aztep')?.value.trim();
          const formKey = document.getElementById('s-aztak')?.value.trim();
          const dep = formDep || Config.azureTtsDeployment;
          const ep  = formEp  || Config.azureTtsEndpoint;
          const key = formKey || Config.azureTtsApiKey || Config.azureApiKey;
          // Persist TTS config to localStorage immediately — don't wait for Save Settings.
          // Without this, Review page _speak() uses stale Config values and hits the wrong endpoint.
          if (formDep) Config.azureTtsDeployment = formDep;
          if (formEp)  Config.azureTtsEndpoint   = formEp;
          if (formKey) Config.azureTtsApiKey     = formKey;
          if (!dep) { this._toast('Enter your Azure TTS Deployment name above first.', 'error'); return; }
          if (!key) { this._toast('Azure API key is not configured.', 'error'); return; }
          this._previewTtsVoice(voice, dep, ep, key);
        });
      });
    }

    if (view === 'review') {
      // Inject answer as HTML (user's own content — safe for personal app)
      const answerEl = document.getElementById('answer-html');
      if (answerEl) {
        const card = this.s.queue[this.s.qIdx];
        answerEl.innerHTML = card?.back
          ? card.back
          : '<em style="color:var(--text2)">No answer added yet.</em>';
        answerEl.addEventListener('dblclick', () => this._enterAnswerEdit());
      }
      document.getElementById('question-card')?.addEventListener('click', () => {
        if (!this.s.answerVisible) {
          this.s.answerVisible = true;
          this._render('review');
          const card = this.s.queue[this.s.qIdx];
          if (card?.back) {
            const tmp = document.createElement('div');
            tmp.innerHTML = card.back;
            this._speak(tmp.textContent || '');
          }
        }
      });
      document.getElementById('audio-toggle')?.addEventListener('click', () => this._toggleAudio());
      document.getElementById('user-answer')?.addEventListener('input', e => { this.s.userAnswer = e.target.value; });
      document.getElementById('mic-btn')?.addEventListener('click', () => this._toggleMic());
      document.getElementById('ai-validate-btn')?.addEventListener('click', () => this._aiValidate());
      document.getElementById('ai-eval-close')?.addEventListener('click', () => {
        this.s.aiEvalText = ''; this.s.aiEvalLoading = false; this._render('review');
      });
      document.getElementById('ai-eval-play')?.addEventListener('click', () => {
        if (this.s.aiEvalText) this._speak(this.s.aiEvalText);
      });
      // Review left-panel collapse toggle
      document.getElementById('review-left-toggle')?.addEventListener('click', () => {
        this.s.reviewLeftCollapsed = !this.s.reviewLeftCollapsed;
        this._render('review');
      });
      // Review left-panel tab switcher
      document.querySelectorAll('[data-review-tab]').forEach(b =>
        b.addEventListener('click', () => {
          this.s.reviewLeftTab = b.dataset.reviewTab;
          this._render('review');
        }));
      // Deck tree (review)
      document.querySelectorAll('[data-dsel]').forEach(b =>
        b.addEventListener('click', () => this._toggleDeck(b.dataset.dsel)));
      document.querySelectorAll('.deck-toggle').forEach(b =>
        b.addEventListener('click', e => { e.stopPropagation(); this._toggleReviewDeckNode(b.dataset.dtoggle); }));
      document.querySelectorAll('.active-deck-rm').forEach(b =>
        b.addEventListener('click', e => { e.stopPropagation(); this._removeDeck(b.dataset.rm); }));
      document.getElementById('clear-decks')?.addEventListener('click', () => {
        this.s.selectedDeck = []; this._buildQueue(); this._render('review'); this._speakCurrent();
      });
      document.getElementById('review-deck-search')?.addEventListener('input', e => {
        this.s.reviewDeckSearch = e.target.value;
        this._render('review');
        const el = document.getElementById('review-deck-search');
        if (el) { el.focus(); el.setSelectionRange(e.target.selectionStart, e.target.selectionStart); }
      });
      document.querySelectorAll('.rating-btn').forEach(b =>
        b.addEventListener('click', () => this._rate(parseInt(b.dataset.r))));
      document.getElementById('go-add')?.addEventListener('click', () => { this.s.editCard=null; this._render('add'); });
      document.getElementById('go-browse')?.addEventListener('click', () => this._render('browse'));
      document.getElementById('restart-btn')?.addEventListener('click', () => { this._buildQueue(); this._render('review'); this._speakCurrent(); });
      document.getElementById('expand-answer')?.addEventListener('click', () => this._toggleAnswerExpand());
      document.getElementById('answer-col-close')?.addEventListener('click', () => {
        this.s.answerVisible = false;
        this.s.answerExpanded = false;
        this._render('review');
      });
      document.getElementById('answer-play')?.addEventListener('click', () => {
        const card = this.s.queue[this.s.qIdx];
        if (card?.back) {
          const tmp = document.createElement('div');
          tmp.innerHTML = card.back;
          this._speak(tmp.textContent || '');
        }
      });

      // ── Tag filter bindings ──
      document.getElementById('tag-search')?.addEventListener('input', e => {
        const sel = e.target.selectionStart;
        this.s.tagSearch = e.target.value;
        this._render('review');
        const el = document.getElementById('tag-search');
        if (el) { el.focus(); el.setSelectionRange(sel, sel); }
      });
      document.getElementById('clear-tags')?.addEventListener('click', () => {
        this.s.selectedTags = []; this._buildQueue(); this._render('review'); this._speakCurrent();
      });
      document.querySelectorAll('.active-tag-rm').forEach(b =>
        b.addEventListener('click', e => { e.stopPropagation(); this._removeReviewTag(b.dataset.rm); }));
      // Tree tag select (replaces flat .tag-filter-item)
      document.querySelectorAll('[data-tsel]').forEach(b =>
        b.addEventListener('click', () => this._toggleTag(b.dataset.tsel)));
      // Tree expand/collapse
      document.querySelectorAll('.tree-toggle').forEach(b =>
        b.addEventListener('click', e => { e.stopPropagation(); this._toggleReviewTreeNode(b.dataset.toggle); }));
      // Reset by tag or group
      document.querySelectorAll('[data-reset-tag]').forEach(b =>
        b.addEventListener('click', e => { e.stopPropagation(); this._resetTagCards(b.dataset.resetTag); }));
      document.querySelectorAll('[data-reset-gid]').forEach(b =>
        b.addEventListener('click', e => { e.stopPropagation(); this._resetGroupCards(b.dataset.resetGid); }));
      document.getElementById('save-tag-group')?.addEventListener('click', () => this._saveTagGroup());
      document.querySelectorAll('.tg-load-btn').forEach(b =>
        b.addEventListener('click', () => this._loadTagGroup(b.dataset.gid)));
      document.querySelectorAll('.tg-del-btn').forEach(b =>
        b.addEventListener('click', () => this._deleteTagGroup(b.dataset.gid)));

      // ── Mobile drawer (review) ──
      document.getElementById('tags-drawer-btn')?.addEventListener('click', () => {
        this.s.tagsDrawerOpen = !this.s.tagsDrawerOpen;
        this._render('review');
      });
      document.getElementById('drawer-backdrop')?.addEventListener('click', () => {
        this.s.tagsDrawerOpen = false;
        this._render('review');
      });

      // ── Editor toolbar bindings (mousedown to keep focus/selection) ──
      document.querySelectorAll('[data-cmd]').forEach(btn => {
        btn.addEventListener('mousedown', e => {
          e.preventDefault();
          document.execCommand(btn.dataset.cmd, false, null);
        });
      });
      document.getElementById('tb-code')?.addEventListener('mousedown', e => { e.preventDefault(); this._insertCode(); });
      document.getElementById('tb-html')?.addEventListener('click', () => this._toggleHtmlView());
      document.getElementById('tb-save')?.addEventListener('click', () => this._saveAnswerEdit());
      document.getElementById('tb-cancel')?.addEventListener('click', () => this._cancelAnswerEdit());
    }

    if (view === 'browse') {
      // Inject answer HTML for selected card
      const bAns = document.getElementById('browse-answer-html');
      if (bAns && this.s.browseSelectedCard) {
        bAns.innerHTML = this.s.browseSelectedCard.back || '<em style="color:var(--text2)">No answer yet.</em>';
      }

      document.getElementById('go-add')?.addEventListener('click', () => { this.s.editCard=null; this._render('add'); });
      document.getElementById('ai-gen-btn')?.addEventListener('click', () => this._aiGenerate());
      document.querySelectorAll('[data-edit]').forEach(b =>
        b.addEventListener('click', e => { e.stopPropagation(); this.s.editCard = this.s.cards.find(c=>c.id===b.dataset.edit)||null; this._render('add'); }));
      document.querySelectorAll('[data-del]').forEach(b =>
        b.addEventListener('click', e => { e.stopPropagation(); this._deleteCard(b.dataset.del); }));

      // Card click → detail preview (on .bci-main child, not checkbox)
      document.querySelectorAll('.bci-main').forEach(b =>
        b.addEventListener('click', () => this._selectBrowseCard(b.dataset.cid)));

      // Multi-select checkboxes
      document.querySelectorAll('.bci-cb').forEach(cb =>
        cb.addEventListener('change', e => { e.stopPropagation(); this._toggleBrowseCardSel(cb.dataset.cbid); }));
      document.getElementById('bulk-assign-deck')?.addEventListener('click', () => this._bulkAssignDeck());
      document.getElementById('bulk-clear-sel')?.addEventListener('click', () => {
        this.s.browseSelectedCards = []; this._render('browse');
      });

      // Browse panel collapse toggles
      document.getElementById('browse-left-toggle')?.addEventListener('click', () => {
        this.s.browseLeftCollapsed = !this.s.browseLeftCollapsed;
        this._render('browse');
      });
      // Tab switcher
      document.querySelectorAll('[data-browse-tab]').forEach(b =>
        b.addEventListener('click', () => {
          this.s.browseLeftTab = b.dataset.browseTab;
          this._render('browse');
        }));

      // Tag tree interactions
      document.getElementById('browse-tag-search')?.addEventListener('input', e => {
        const sel = e.target.selectionStart;
        this.s.browseTagSearch = e.target.value;
        this._render('browse');
        const el = document.getElementById('browse-tag-search');
        if (el) { el.focus(); el.setSelectionRange(sel, sel); }
      });
      document.getElementById('browse-search')?.addEventListener('input', e => {
        const sel = e.target.selectionStart;
        this.s.browseSearch = e.target.value;
        this._render('browse');
        const el = document.getElementById('browse-search');
        if (el) { el.focus(); el.setSelectionRange(sel, sel); }
      });
      document.querySelectorAll('.tree-toggle').forEach(b =>
        b.addEventListener('click', e => { e.stopPropagation(); this._toggleTreeNode(b.dataset.toggle); }));
      document.querySelectorAll('[data-tsel]').forEach(b =>
        b.addEventListener('click', () => this._toggleBrowseTag(b.dataset.tsel)));
      document.querySelectorAll('.active-tag-rm').forEach(b =>
        b.addEventListener('click', e => { e.stopPropagation(); this._removeBrowseTag(b.dataset.rm); }));
      document.getElementById('clear-browse-tags')?.addEventListener('click', () => {
        this.s.browseFilterTags = []; this._render('browse');
      });

      // Deck tree interactions (browse)
      document.getElementById('browse-deck-search')?.addEventListener('input', e => {
        const sel = e.target.selectionStart;
        this.s.browseDeckSearch = e.target.value;
        this._render('browse');
        const el = document.getElementById('browse-deck-search');
        if (el) { el.focus(); el.setSelectionRange(sel, sel); }
      });
      document.querySelectorAll('.deck-toggle').forEach(b =>
        b.addEventListener('click', e => { e.stopPropagation(); this._toggleDeckTreeNode(b.dataset.dtoggle); }));
      document.querySelectorAll('[data-dsel]').forEach(b =>
        b.addEventListener('click', () => this._toggleBrowseDeck(b.dataset.dsel)));
      document.querySelectorAll('.active-deck-rm').forEach(b =>
        b.addEventListener('click', e => { e.stopPropagation(); this._removeBrowseDeck(b.dataset.rm); }));
      document.getElementById('clear-browse-decks')?.addEventListener('click', () => {
        this.s.browseDeckFilter = []; this._render('browse');
      });

      // ── Mobile drawer (browse) ──
      document.getElementById('browse-filter-btn')?.addEventListener('click', () => {
        this.s.browseTagsDrawerOpen = !this.s.browseTagsDrawerOpen;
        this._render('browse');
      });
      document.getElementById('browse-backdrop')?.addEventListener('click', () => {
        this.s.browseTagsDrawerOpen = false;
        this._render('browse');
      });
      document.getElementById('browse-detail-close')?.addEventListener('click', () => {
        this.s.browseSelectedCard = null;
        this._render('browse');
      });
    }

    if (view === 'decks') {
      document.getElementById('deck-new-btn')?.addEventListener('click', () => {
        this.s.decksShowNewForm = true;
        this._render('decks');
        setTimeout(() => document.getElementById('deck-new-input')?.focus(), 0);
      });

      document.getElementById('deck-new-input')?.addEventListener('input', e => {
        this.s.decksNewInput = e.target.value;
      });

      // New-deck upload: button triggers the file input programmatically
      document.getElementById('deck-new-upload-btn')?.addEventListener('click', () => {
        const inputEl = document.getElementById('deck-new-input');
        const name = (inputEl ? inputEl.value : this.s.decksNewInput).trim();
        if (!name) { this._toast('Enter a deck name first, then click Upload.', 'error'); return; }
        document.getElementById('deck-new-file')?.click();
      });

      document.getElementById('deck-new-file')?.addEventListener('change', e => {
        const file = e.target.files[0];
        if (!file) return;
        const inputEl = document.getElementById('deck-new-input');
        const name = (inputEl ? inputEl.value : this.s.decksNewInput).trim();
        if (!name) { this._toast('Enter a deck name first, then click Upload.', 'error'); return; }
        this.s.decksShowNewForm = false;
        this.s.decksNewInput = '';
        this._uploadDeckCards(name, file);
      });

      document.getElementById('deck-new-add-manual')?.addEventListener('click', () => {
        const name = this.s.decksNewInput.trim();
        this.s.decksShowNewForm = false;
        this.s.decksNewInput = '';
        this.s.editCard = name ? { deck: name } : null;
        this._render('add');
      });

      document.getElementById('deck-create-cancel')?.addEventListener('click', () => {
        this.s.decksShowNewForm = false;
        this.s.decksNewInput = '';
        this._render('decks');
      });

      document.querySelectorAll('[data-dmtoggle]').forEach(b =>
        b.addEventListener('click', e => {
          e.stopPropagation();
          this._toggleDeckManagerNode(b.dataset.dmtoggle);
        }));

      document.querySelectorAll('[data-dmnav]').forEach(b =>
        b.addEventListener('click', () => {
          this.s.browseDeckFilter = [b.dataset.dmnav];
          this.s.browseLeftTab = 'decks';
          this._render('browse');
        }));

      // Upload button → trigger sibling file input
      document.querySelectorAll('[data-uploadbtn]').forEach(btn =>
        btn.addEventListener('click', e => {
          e.stopPropagation();
          const deck = btn.dataset.uploadbtn;
          const inp = btn.nextElementSibling; // file input is immediately after the button
          if (inp) inp.click();
        }));

      document.querySelectorAll('.deck-file-input').forEach(inp =>
        inp.addEventListener('change', e => {
          const file = e.target.files[0];
          if (!file) return;
          this._uploadDeckCards(inp.dataset.deck, file);
          e.target.value = '';
        }));

      document.querySelectorAll('[data-dmrename]').forEach(b =>
        b.addEventListener('click', e => {
          e.stopPropagation();
          this._renameDeck(b.dataset.dmrename);
        }));
    }

    if (view === 'add') {
      document.getElementById('save-btn')?.addEventListener('click', e => this._saveCard(e.target.dataset.id));
      document.getElementById('cancel-btn')?.addEventListener('click', () => { this.s.editCard=null; this._render('browse'); });
    }

    if (view === 'import') {
      let importFile = null;

      const fileInput = document.getElementById('import-file');
      const dropZone  = document.getElementById('import-drop-zone');
      const fileName  = document.getElementById('import-file-name');
      const startBtn  = document.getElementById('import-start-btn');

      const setFile = f => {
        importFile = f;
        fileName.textContent = f ? f.name : 'No file selected';
        startBtn.disabled = !f;
      };

      document.getElementById('import-file-link')?.addEventListener('click', () => fileInput.click());
      dropZone?.addEventListener('click', () => fileInput.click());
      fileInput?.addEventListener('change', () => setFile(fileInput.files[0] || null));

      dropZone?.addEventListener('dragover', e => { e.preventDefault(); dropZone.style.borderColor = 'var(--primary)'; });
      dropZone?.addEventListener('dragleave', () => { dropZone.style.borderColor = ''; });
      dropZone?.addEventListener('drop', e => {
        e.preventDefault();
        dropZone.style.borderColor = '';
        setFile(e.dataTransfer.files[0] || null);
      });

      document.getElementById('import-cancel-btn')?.addEventListener('click', () => this._render('browse'));
      document.getElementById('import-go-browse')?.addEventListener('click', () => this._render('browse'));
      document.getElementById('import-again-btn')?.addEventListener('click', () => this._render('import'));

      startBtn?.addEventListener('click', async () => {
        if (!importFile) return;
        const deckFilter = document.getElementById('import-deck-filter').value.trim();
        const mediaMode  = document.getElementById('import-media-mode').value;

        // Show progress UI
        document.getElementById('import-progress').style.display = 'block';
        document.getElementById('import-result').style.display = 'none';
        startBtn.disabled = true;

        const log   = document.getElementById('import-log');
        const bar   = document.getElementById('import-progress-bar');
        const msg   = document.getElementById('import-status-msg');
        const addLog = (line) => { log.textContent += line + '\n'; log.scrollTop = log.scrollHeight; };

        try {
          // Dynamically load the import worker script
          if (!window.AnkiImporter) {
            await new Promise((res, rej) => {
              const s = document.createElement('script');
              s.src = 'js/import.js';
              s.onload = res; s.onerror = rej;
              document.head.appendChild(s);
            });
          }

          msg.textContent = 'Reading package…';
          bar.style.width = '5%';

          const result = await AnkiImporter.run({
            file: importFile,
            deckFilter,
            mediaMode,
            token: Auth.token,
            onProgress: (pct, text) => {
              bar.style.width = pct + '%';
              msg.textContent = text;
              addLog(text);
            }
          });

          bar.style.width = '100%';
          msg.textContent = 'Done!';

          // Reload cards/progress from sheet
          [this.s.cards, this.s.progress] = await Promise.all([Sheets.loadCards(), Sheets.loadProgress()]);
          await Media.loadMap();
          this._buildQueue();

          document.getElementById('import-result').style.display = 'block';
          document.getElementById('import-result-msg').innerHTML =
            `<span style="color:var(--success)">✓</span> Imported ${result.cards} card(s)` +
            (result.mediaUploaded ? `, ${result.mediaUploaded} media file(s) uploaded.` : '.');
        } catch (e) {
          addLog('ERROR: ' + e.message);
          msg.textContent = 'Import failed — see log below.';
          bar.style.width = '100%';
          bar.style.background = 'var(--danger)';
          startBtn.disabled = false;
        }
      });
    }
  },

  // ── Keyboard shortcuts ───────────────────────────────────────────────────────
  _onKey(e) {
    if (this.s.view !== 'review') return;
    const tag = document.activeElement?.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA') return;
    if (['1','2','3','4'].includes(e.key)) { this._rate(+e.key); }
  },

  // ── Actions ──────────────────────────────────────────────────────────────────

  async _rate(grade) {
    const card    = this.s.queue[this.s.qIdx];
    const oldProg = this.s.progress[card.id] || SM2.defaultProgress();
    const newProg = { ...SM2.review(oldProg, grade), card_id: card.id };

    this._stopListening();
    this.s.progress[card.id] = newProg;
    this.s.qIdx++;
    this.s.sessionReviewed++;
    if (grade >= 3) this.s.sessionCorrect++;
    this.s.flipped           = false;
    this.s.aiEvalText        = '';
    this.s.aiEvalLoading     = false;
    this.s.answerColCollapsed = false;
    this.s.userAnswer     = '';
    this.s.answerVisible  = false;
    this._render('review');
    this._speakCurrent();

    try { await Sheets.saveProgress(newProg); }
    catch (e) { this._toast('Could not save progress: ' + e.message, 'error'); }
  },

  // ── Browse actions ───────────────────────────────────────────────────────────

  _toggleBrowseTag(path) {
    const idx = this.s.browseFilterTags.indexOf(path);
    if (idx >= 0) this.s.browseFilterTags.splice(idx, 1);
    else this.s.browseFilterTags.push(path);
    this._render('browse');
  },

  _removeBrowseTag(path) {
    this.s.browseFilterTags = this.s.browseFilterTags.filter(t => t !== path);
    this._render('browse');
  },

  _toggleTreeNode(path) {
    this.s.tagTreeExpanded[path] = !this.s.tagTreeExpanded[path];
    this._render('browse');
  },

  // ── Deck tree expand/collapse ─────────────────────────────────────────────────
  _toggleDeckTreeNode(path) {
    this.s.deckTreeExpanded[path] = !this.s.deckTreeExpanded[path];
    this._render('browse');
  },

  _toggleReviewDeckNode(path) {
    this.s.reviewDeckTreeExpanded[path] = !this.s.reviewDeckTreeExpanded[path];
    this._render('review');
  },

  // ── Browse deck filter ────────────────────────────────────────────────────────
  _toggleBrowseDeck(path) {
    const idx = this.s.browseDeckFilter.indexOf(path);
    if (idx >= 0) this.s.browseDeckFilter.splice(idx, 1);
    else this.s.browseDeckFilter.push(path);
    this._render('browse');
  },

  _removeBrowseDeck(path) {
    this.s.browseDeckFilter = this.s.browseDeckFilter.filter(d => d !== path);
    this._render('browse');
  },

  // ── Review deck filter ────────────────────────────────────────────────────────
  _toggleDeck(path) {
    const idx = this.s.selectedDeck.indexOf(path);
    if (idx >= 0) this.s.selectedDeck.splice(idx, 1);
    else this.s.selectedDeck.push(path);
    this._buildQueue();
    this._render('review');
    this._speakCurrent();
  },

  _removeDeck(path) {
    this.s.selectedDeck = this.s.selectedDeck.filter(d => d !== path);
    this._buildQueue();
    this._render('review');
    this._speakCurrent();
  },

  // ── Browse multi-select ───────────────────────────────────────────────────────
  _toggleBrowseCardSel(id) {
    const idx = this.s.browseSelectedCards.indexOf(id);
    if (idx >= 0) this.s.browseSelectedCards.splice(idx, 1);
    else this.s.browseSelectedCards.push(id);
    this._render('browse');
  },

  async _bulkAssignDeck() {
    const allDecks = [...new Set(this.s.cards.map(c => c.deck || 'Default'))].sort();
    const hint = allDecks.length ? `\n\nExisting decks:\n${allDecks.join(', ')}` : '';
    const deck = prompt(`Move ${this.s.browseSelectedCards.length} card(s) to deck:${hint}`, 'Default');
    if (!deck?.trim()) return;
    const target = deck.trim();
    const ids = new Set(this.s.browseSelectedCards);
    const toUpdate = this.s.cards.filter(c => ids.has(c.id));
    this._toast(`Saving ${toUpdate.length} card(s)…`);
    let failed = 0;
    for (const card of toUpdate) {
      card.deck = target;
      try { await Sheets.updateCard(card); }
      catch { failed++; }
    }
    this.s.browseSelectedCards = [];
    this._buildQueue();
    this._render('browse');
    if (failed) this._toast(`Done — ${failed} save(s) failed, try again.`, 'error');
    else this._toast(`Moved ${toUpdate.length} card(s) to "${target}".`, 'success');
  },

  _selectBrowseCard(cardId) {
    this.s.browseSelectedCard = this.s.browseSelectedCard?.id === cardId
      ? null
      : (this.s.cards.find(c => c.id === cardId) || null);
    this._render('browse');
  },

  // ── Tag filter actions ───────────────────────────────────────────────────────

  _toggleTag(tag) {
    const idx = this.s.selectedTags.indexOf(tag);
    if (idx >= 0) this.s.selectedTags.splice(idx, 1);
    else this.s.selectedTags.push(tag);
    this._buildQueue();
    this._render('review');
    this._speakCurrent();
  },

  _removeReviewTag(path) {
    this.s.selectedTags = this.s.selectedTags.filter(t => t !== path);
    this._buildQueue();
    this._render('review');
    this._speakCurrent();
  },

  async _saveTagGroup() {
    const name = prompt('Name for this tag group:');
    if (!name?.trim()) return;
    try {
      const group = await Sheets.saveTagGroup({ name: name.trim(), tags: this.s.selectedTags.join(' :: ') });
      this.s.tagGroups.push(group);
      this._render('review');
      this._toast('Tag group saved!', 'success');
    } catch (e) {
      this._toast('Save failed: ' + e.message, 'error');
    }
  },

  async _deleteTagGroup(gid) {
    if (!confirm('Delete this tag group?')) return;
    try {
      await Sheets.deleteTagGroup(gid);
      this.s.tagGroups = this.s.tagGroups.filter(g => g.id !== gid);
      this._render('review');
      this._toast('Group deleted.', 'success');
    } catch (e) {
      this._toast('Delete failed: ' + e.message, 'error');
    }
  },

  _loadTagGroup(gid) {
    const group = this.s.tagGroups.find(g => g.id === gid);
    if (!group) return;
    this.s.selectedTags = parseTags(group.tags);
    this._buildQueue();
    this._render('review');
    this._speakCurrent();
  },

  _toggleAnswerExpand() {
    this.s.answerExpanded = !this.s.answerExpanded;
    this._render('review');
  },

  _toggleReviewTreeNode(path) {
    this.s.reviewTreeExpanded[path] = !this.s.reviewTreeExpanded[path];
    this._render('review');
  },

  // ── Reset actions ────────────────────────────────────────────────────────────

  async _resetTagCards(tagPath) {
    const affected = this.s.cards.filter(c => {
      const ct = parseTags(c.tags || '');
      return ct.some(t => t === tagPath || t.startsWith(tagPath + '::'));
    });
    if (!affected.length) { this._toast('No cards found for this tag.'); return; }
    if (!confirm(`Reset ${affected.length} card(s) tagged "${tagDisplayName(tagPath)}"?\nThey will be due for review immediately.`)) return;
    await this._resetCards(affected.map(c => c.id));
  },

  async _resetGroupCards(gid) {
    const group = this.s.tagGroups.find(g => g.id === gid);
    if (!group) return;
    const gTags = parseTags(group.tags);
    const affected = this.s.cards.filter(c => {
      const ct = parseTags(c.tags || '');
      return gTags.some(sel => ct.some(t => t === sel || t.startsWith(sel + '::')));
    });
    if (!affected.length) { this._toast('No cards found in this group.'); return; }
    if (!confirm(`Reset ${affected.length} card(s) in group "${group.name}"?\nThey will be due for review immediately.`)) return;
    await this._resetCards(affected.map(c => c.id));
  },

  async _resetCards(cardIds) {
    const today = new Date().toISOString().split('T')[0];
    this._toast(`Resetting ${cardIds.length} card(s)…`);
    let failed = 0;
    for (const id of cardIds) {
      const prog = { card_id: id, easeFactor: 2.5, interval: 0, repetitions: 0, dueDate: today, lastReview: '' };
      this.s.progress[id] = prog;
      try { await Sheets.saveProgress(prog); }
      catch { failed++; }
    }
    this._buildQueue();
    this._render('review');
    this._speakCurrent();
    if (failed) this._toast(`Reset done — ${failed} save(s) failed, refresh may fix it.`, 'error');
    else this._toast(`Reset ${cardIds.length} card(s) — they're due now!`, 'success');
  },

  async _aiValidate() {
    const card = this.s.queue[this.s.qIdx];
    if (!card) return;
    this.s.aiEvalLoading = true;
    this.s.aiEvalText    = '';
    this._render('review');
    const hasAnswer = !!this.s.userAnswer.trim();
    this.s.aiEvalText = hasAnswer
      ? await LLM.validate(card.front, card.back, this.s.userAnswer)
      : await LLM.explain(card.front, card.back);
    this.s.aiEvalLoading = false;
    this._speak(this.s.aiEvalText);
    this._render('review');
  },


  async _signIn() {
    try {
      this._render('loading');
      await Auth.signIn();
    } catch (e) {
      this._toast('Sign-in failed: ' + e.message, 'error');
      this._render('login');
      return;
    }
    try {
      await this._load();
      this._render('review');
      this._speakCurrent();
      // Nudge the user if AI is not yet usable
      if (!Config.azureApiKey) {
        setTimeout(() => this._toast('AI features need an Azure API key — add it to your Sheet\'s Settings tab (key: azureApiKey) or enter it on the Settings page.'), 600);
      }
    } catch (e) {
      this._toast(e.message, 'error');
      this._render('settings');
    }
  },

  _signOut() {
    Auth.signOut();
    this._render('login');
  },

  _populateVoices() {
    const container = document.getElementById('voice-list');
    if (!container || !window.speechSynthesis) {
      if (container) container.innerHTML = '<div class="voice-loading">Speech not supported in this browser.</div>';
      return;
    }
    const render = () => {
      const all = window.speechSynthesis.getVoices();
      const voices = all
        .filter(v => v.lang.startsWith('en'))
        .sort((a, b) => {
          // HD voices first, then local, then alphabetical
          const aHD = a.name.includes('Enhanced') || a.name.includes('Premium');
          const bHD = b.name.includes('Enhanced') || b.name.includes('Premium');
          if (aHD !== bHD) return aHD ? -1 : 1;
          if (a.localService !== b.localService) return a.localService ? -1 : 1;
          return a.name.localeCompare(b.name);
        });
      if (!voices.length) {
        container.innerHTML = '<div class="voice-loading">No English voices found.</div>';
        return;
      }
      const cur = Config.voiceName;
      container.innerHTML = voices.map(v => {
        const hd = v.name.includes('Enhanced') || v.name.includes('Premium') || v.name.includes('Neural');
        return `
          <div class="voice-item${v.name === cur ? ' sel' : ''}" data-voice="${h(v.name)}">
            <div class="voice-info">
              <span class="voice-name">${h(v.name)}</span>
              <span class="voice-lang">${h(v.lang)}</span>
              ${hd ? '<span class="voice-badge">HD</span>' : ''}
            </div>
            <button class="voice-sample-btn" data-sample="${h(v.name)}" title="Preview">▶</button>
          </div>`;
      }).join('');

      container.querySelectorAll('.voice-item').forEach(el => {
        el.addEventListener('click', () => {
          Config.voiceName = el.dataset.voice;
          container.querySelectorAll('.voice-item').forEach(e => e.classList.remove('sel'));
          el.classList.add('sel');
          this._sampleVoice(el.dataset.voice);
        });
      });
      container.querySelectorAll('.voice-sample-btn').forEach(btn => {
        btn.addEventListener('click', e => { e.stopPropagation(); this._sampleVoice(btn.dataset.sample); });
      });
    };

    const existing = window.speechSynthesis.getVoices();
    if (existing.length) render();
    else window.speechSynthesis.onvoiceschanged = render;
  },

  _sampleVoice(name) {
    if (!window.speechSynthesis) return;
    window.speechSynthesis.cancel();
    const voice = window.speechSynthesis.getVoices().find(v => v.name === name);
    const u = new SpeechSynthesisUtterance('Hello! This is how I sound. I will read your flashcard answers in this voice.');
    u.rate  = 0.88;
    u.pitch = 1.0;
    if (voice) u.voice = voice;
    window.speechSynthesis.speak(u);
  },

  async _reloadFromSheet() {
    const btn = document.getElementById('reload-from-sheet');
    if (btn) btn.disabled = true;
    try {
      const settings = await Sheets.loadSettings();
      this._applySheetSettings(settings);
      this._toast('Settings reloaded from Sheet.', 'success');
      this._render('settings');  // re-render so inputs show updated values
    } catch (e) {
      this._toast('Reload failed: ' + e.message, 'error');
      if (btn) btn.disabled = false;
    }
  },

  async _saveSettings() {
    const gci = document.getElementById('s-gci')?.value.trim();
    const gak = document.getElementById('s-gak')?.value.trim();
    const sid = document.getElementById('s-sid')?.value.trim();
    const azk    = document.getElementById('s-azk')?.value.trim();
    const aze    = document.getElementById('s-aze')?.value.trim();
    const azd    = document.getElementById('s-azd')?.value.trim();
    const azv    = document.getElementById('s-azv')?.value.trim();
    const aztep  = document.getElementById('s-aztep')?.value.trim();
    const aztak  = document.getElementById('s-aztak')?.value.trim();
    const aztts  = document.getElementById('s-aztts')?.value.trim();
    const azttsv = document.getElementById('s-azttsv')?.value.trim();

    if (!gci) {
      this._toast('Google Client ID is required.', 'error');
      return;
    }
    if (!sid) {
      this._toast('Google Sheet ID is required.', 'error');
      return;
    }
    Config.googleClientId = gci;
    Config.googleApiKey   = gak;
    Config.sheetId        = sid;
    if (azk)   Config.azureApiKey        = azk;
    if (aze)   Config.azureEndpoint      = aze;
    if (azd)   Config.azureDeployment    = azd;
    if (azv)   Config.azureApiVersion    = azv;
    Config.azureTtsEndpoint   = aztep || '';
    if (aztak) Config.azureTtsApiKey   = aztak;
    Config.azureTtsDeployment = aztts || '';
    if (azttsv) Config.azureTtsApiVersion = azttsv;

    this._toast('Settings saved!', 'success');
    if (this.s.view === 'setup') setTimeout(() => this._render('login'), 700);

    // Persist to Settings sheet in background (only when signed in)
    if (Auth.isSignedIn()) {
      const toSave = { theme: Config.theme, ttsVoice: Config.ttsVoice };
      if (azk)    toSave.azureApiKey         = azk;
      if (aze)    toSave.azureEndpoint       = aze;
      if (azd)    toSave.azureDeployment     = azd;
      if (azv)    toSave.azureApiVersion     = azv;
      if (aztep)  toSave.azureTtsEndpoint   = aztep;
      if (aztak)  toSave.azureTtsApiKey     = aztak;
      if (aztts)  toSave.azureTtsDeployment = aztts;
      if (azttsv) toSave.azureTtsApiVersion = azttsv;
      for (const [k, v] of Object.entries(toSave)) {
        Sheets.saveSetting(k, v).catch(() => {});
      }
    }
  },

  // ── Shared card insertion — used by Add Card form, file upload, and AI generate ──
  async _insertCard(cardData) {
    const card = await Sheets.saveCard({
      front: (cardData.front || '').trim(),
      back:  (cardData.back  || '').trim(),
      tags:  (cardData.tags  || '').trim(),
      notes: (cardData.notes || '').trim(),
      deck:  (cardData.deck  || 'Default').trim()
    });
    this.s.cards.push(card);
    return card;
  },

  async _saveCard(id) {
    const front = document.getElementById('f-front')?.value.trim();
    const back  = document.getElementById('f-back')?.value.trim();
    const deck  = (document.getElementById('f-deck')?.value.trim() || 'Default');
    const tags  = document.getElementById('f-tags')?.value.trim();
    const notes = document.getElementById('f-notes')?.value.trim();

    if (!front) { this._toast('Question is required.', 'error'); return; }

    const btn = document.getElementById('save-btn');
    if (btn) btn.disabled = true;

    try {
      if (id) {
        const card = this.s.cards.find(c => c.id === id);
        Object.assign(card, { front, back, deck, tags, notes });
        await Sheets.updateCard(card);
        this._toast('Card updated.', 'success');
      } else {
        await this._insertCard({ front, back, deck, tags, notes });
        this._toast('Card added!', 'success');
      }
      this.s.editCard = null;
      this._buildQueue();
      this._render('browse');
    } catch (e) {
      this._toast('Save failed: ' + e.message, 'error');
      if (btn) btn.disabled = false;
    }
  },

  async _deleteCard(id) {
    if (!confirm('Delete this card? It will be cleared from your Sheet.')) return;
    try {
      await Sheets.deleteCard(id);
      this.s.cards = this.s.cards.filter(c => c.id !== id);
      delete this.s.progress[id];
      if (this.s.browseSelectedCard?.id === id) this.s.browseSelectedCard = null;
      this._buildQueue();
      this._render('browse');
      this._toast('Card deleted.', 'success');
    } catch (e) {
      this._toast('Delete failed: ' + e.message, 'error');
    }
  },

  async _aiGenerate() {
    const topic = prompt('Topic for AI-generated cards?\n(e.g. "React hooks", "SQL joins", "System design basics")');
    if (!topic?.trim()) return;
    this._toast('Generating cards…');
    try {
      const cards = await LLM.generateCards(topic.trim(), 5);
      for (const c of cards) {
        await this._insertCard(c);
      }
      this._buildQueue();
      this._render('browse');
      this._toast(`Added ${cards.length} new cards!`, 'success');
    } catch (e) {
      this._toast('Generation failed: ' + e.message, 'error');
    }
  },

  // ── Deck manager actions ─────────────────────────────────────────────────────

  _toggleDeckManagerNode(path) {
    this.s.deckManagerExpanded[path] = !this.s.deckManagerExpanded[path];
    this._render('decks');
  },

  async _renameDeck(oldPath) {
    const current = tagDisplayName(oldPath);
    const input = prompt(`Rename deck "${current}" to:\n(Use :: for hierarchy, e.g. Java::Collections)`, current);
    if (!input?.trim()) return;
    // Accept either › or :: as separator in the typed name
    const newPath = input.trim().replace(/ › /g, '::').replace(/ > /g, '::');
    if (newPath === oldPath) return;

    const affected = this.s.cards.filter(c => {
      const d = (c.deck || 'Default').trim();
      return d === oldPath || d.startsWith(oldPath + '::');
    });
    if (!affected.length) { this._toast('No cards found in this deck.'); return; }
    if (!confirm(`Rename "${current}" → "${tagDisplayName(newPath)}"?\n${affected.length} card(s) will be updated.`)) return;

    this._toast(`Renaming ${affected.length} card(s)…`);
    let failed = 0;
    for (const card of affected) {
      const oldDeck = (card.deck || 'Default').trim();
      card.deck = oldDeck === oldPath
        ? newPath
        : newPath + oldDeck.slice(oldPath.length);
      try { await Sheets.updateCard(card); }
      catch (err) { console.error('updateCard failed:', err); failed++; }
    }
    this._buildQueue();
    this._render('decks');
    if (failed) this._toast(`Renamed — ${failed} save(s) failed.`, 'error');
    else this._toast(`Renamed to "${tagDisplayName(newPath)}"!`, 'success');
  },

  async _uploadDeckCards(deckPath, file) {
    let text;
    try { text = await file.text(); }
    catch (err) { this._toast('Could not read file.', 'error'); return; }

    const delimiter = this._detectDelimiter(text, file.name);
    const rows = this._parseCSVTSV(text, delimiter);
    if (!rows.length) { this._toast('No rows found in file — check that first row has headers (front, back, tags, notes).', 'error'); return; }

    this._toast(`Uploading ${rows.length} card(s) to "${tagDisplayName(deckPath)}"…`);
    let added = 0, failed = 0;
    for (const row of rows) {
      const front = (row.front || row.question || '').trim();
      const back  = (row.back  || row.answer   || '').trim();
      if (!front) continue;
      try {
        await this._insertCard({
          front,
          back,
          tags:  (row.tags  || row.tag   || '').trim(),
          notes: (row.notes || row.note  || '').trim(),
          deck:  (row.deck  || '').trim() || deckPath
        });
        added++;
      } catch (err) { console.error('insertCard failed:', err); failed++; }
    }
    this._buildQueue();
    this._render('decks');
    if (failed) this._toast(`Uploaded ${added} card(s) — ${failed} failed.`, 'error');
    else this._toast(`Uploaded ${added} card(s) to "${tagDisplayName(deckPath)}"!`, 'success');
  },

  // Detect the delimiter used in a CSV/TSV/SSV file
  _detectDelimiter(text, filename) {
    const firstLine = (text.split(/\r?\n/).find(l => l.trim()) || '').trim();
    if (!firstLine) return ',';
    // Prefer tab, then comma, then semicolon, then pipe, then space
    if (firstLine.includes('\t'))   return '\t';
    if (firstLine.includes(','))    return ',';
    if (firstLine.includes(';'))    return ';';
    if (firstLine.includes('|'))    return '|';
    return ' ';
  },

  _parseCSVTSV(text, delimiter = ',') {
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    if (!lines.length) return [];

    const firstVals = this._splitCsvRow(lines[0], delimiter)
      .map(v => v.trim().replace(/^["']|["']$/g, ''));

    // Detect whether first row is a header row
    const knownHdrs = new Set(['front','back','question','answer','tags','tag','notes','note','deck','id']);
    const hasHeader = firstVals.some(v => knownHdrs.has(v.toLowerCase()));

    let headers, dataStart;
    if (hasHeader) {
      headers = firstVals.map(v => v.toLowerCase());
      dataStart = 1;
    } else {
      // No header — infer column mapping from shape of first row
      headers = this._inferHeaders(firstVals);
      dataStart = 0;
    }

    const rows = [];
    for (let i = dataStart; i < lines.length; i++) {
      const vals = this._splitCsvRow(lines[i], delimiter);
      if (vals.every(v => !v.trim())) continue;
      const row = {};
      headers.forEach((hdr, idx) => {
        if (hdr) row[hdr] = (vals[idx] || '').replace(/^["']|["']$/g, '').trim();
      });
      rows.push(row);
    }
    return rows;
  },

  // Infer column names when the file has no header row.
  // Handles formats: num|id|front|back|tags  or  front|back|tags  etc.
  _inferHeaders(firstVals) {
    const n = firstVals.length;
    const uuidRe = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    const col0Num  = /^\d+$/.test(firstVals[0] || '');
    const col1Uuid = n > 1 && uuidRe.test(firstVals[1] || '');

    if (col0Num && col1Uuid) {
      // Format: row_num | uuid | front | back | tags | [notes] | [deck]
      return [null, null, 'front', 'back', 'tags', 'notes', 'deck'].slice(0, n);
    }
    if (col0Num) {
      // Format: row_num | front | back | [tags] | [notes]
      return [null, 'front', 'back', 'tags', 'notes'].slice(0, n);
    }
    // Generic positional
    return ['front', 'back', 'tags', 'notes', 'deck'].slice(0, n);
  },

  _splitCsvRow(line, delimiter) {
    const result = [];
    let inQuote = false, cur = '';
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        inQuote = !inQuote;
      } else if (ch === delimiter && !inQuote) {
        result.push(cur); cur = '';
      } else {
        cur += ch;
      }
    }
    result.push(cur);
    return result;
  },

  // ── Answer rich-text editor ──────────────────────────────────────────────────
  _enterAnswerEdit() {
    const el = document.getElementById('answer-html');
    const toolbar = document.getElementById('editor-toolbar');
    if (!el || el.contentEditable === 'true') return;
    el.contentEditable = 'true';
    el.classList.add('editing');
    toolbar.style.display = 'flex';
    el.focus();
    // Move cursor to end
    const range = document.createRange();
    range.selectNodeContents(el);
    range.collapse(false);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  },

  _insertCode() {
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    const text = range.toString();
    if (!text) return;
    if (text.includes('\n')) {
      document.execCommand('formatBlock', false, 'pre');
    } else {
      const code = document.createElement('code');
      code.textContent = text;
      range.deleteContents();
      range.insertNode(code);
      const next = document.createRange();
      next.setStartAfter(code);
      next.collapse(true);
      sel.removeAllRanges();
      sel.addRange(next);
    }
  },

  _toggleHtmlView() {
    const el = document.getElementById('answer-html');
    const existing = document.getElementById('answer-raw-ta');
    const btn = document.getElementById('tb-html');
    if (existing) {
      // Back to rich text
      el.innerHTML = existing.value;
      existing.remove();
      el.style.display = '';
      btn.textContent = 'HTML';
      el.focus();
    } else {
      // Show raw HTML textarea
      const ta = document.createElement('textarea');
      ta.id = 'answer-raw-ta';
      ta.className = 'answer-raw-ta';
      ta.value = el.innerHTML;
      el.style.display = 'none';
      el.after(ta);
      ta.focus();
      btn.textContent = 'Rich Text';
    }
  },

  async _saveAnswerEdit() {
    const el = document.getElementById('answer-html');
    const rawTa = document.getElementById('answer-raw-ta');
    // Get content from whichever view is active
    let html = rawTa ? rawTa.value : el.innerHTML;
    // Clean up raw-html textarea if open
    if (rawTa) { el.innerHTML = html; rawTa.remove(); el.style.display = ''; }
    // Exit edit mode
    el.contentEditable = 'false';
    el.classList.remove('editing');
    document.getElementById('editor-toolbar').style.display = 'none';
    document.getElementById('tb-html').textContent = 'HTML';
    // Persist to state + Sheets
    const card = this.s.queue[this.s.qIdx];
    if (!card) return;
    card.back = html;
    // Also sync to the cards array
    const master = this.s.cards.find(c => c.id === card.id);
    if (master) master.back = html;
    try {
      await Sheets.updateCard(card);
      this._toast('Answer saved!', 'success');
    } catch (e) {
      this._toast('Save failed: ' + e.message, 'error');
    }
  },

  _cancelAnswerEdit() {
    const el = document.getElementById('answer-html');
    const rawTa = document.getElementById('answer-raw-ta');
    if (rawTa) { rawTa.remove(); el.style.display = ''; }
    el.contentEditable = 'false';
    el.classList.remove('editing');
    document.getElementById('editor-toolbar').style.display = 'none';
    document.getElementById('tb-html').textContent = 'HTML';
    // Restore original
    const card = this.s.queue[this.s.qIdx];
    el.innerHTML = card?.back || '<em style="color:var(--text2)">No answer added yet.</em>';
  },

  // ── Audio (Web Speech API) ───────────────────────────────────────────────────
  _speak(text) {
    if (!this.s.audioOn || !text.trim()) return;
    if (Config.azureTtsDeployment && (Config.azureTtsApiKey || Config.azureApiKey)) {
      this._speakAzure(text);
    } else {
      this._speakBrowser(text);
    }
  },

  async _speakAzure(text) {
    if (this._ttsAudio) { this._ttsAudio.pause(); URL.revokeObjectURL(this._ttsAudio.src); this._ttsAudio = null; }
    try {
      const url = await LLM.tts(text);
      const audio = new Audio(url);
      this._ttsAudio = audio;
      audio.onended = () => { URL.revokeObjectURL(url); this._ttsAudio = null; };
      await audio.play();
    } catch (e) {
      this._toast('Azure TTS error: ' + e.message, 'error');
      this._speakBrowser(text);
    }
  },

  async _previewTtsVoice(voice, dep, ep, key) {
    if (this._ttsAudio) { this._ttsAudio.pause(); URL.revokeObjectURL(this._ttsAudio.src); this._ttsAudio = null; }
    try {
      const url = await LLM.tts(`Hi, I'm ${voice}. This is how I sound reading your flashcards.`, dep, voice, ep, key);
      const audio = new Audio(url);
      this._ttsAudio = audio;
      audio.onended = () => { URL.revokeObjectURL(url); this._ttsAudio = null; };
      await audio.play();
    } catch (e) {
      this._toast('TTS preview failed: ' + e.message, 'error');
    }
  },

  _speakBrowser(text) {
    if (!window.speechSynthesis) return;
    window.speechSynthesis.cancel();
    const voices = window.speechSynthesis.getVoices();
    const voice  = Config.voiceName ? voices.find(v => v.name === Config.voiceName) : null;
    const makeU = (t) => {
      const u = new SpeechSynthesisUtterance(t);
      u.rate  = 0.88;
      u.pitch = 1.0;
      if (voice) u.voice = voice;
      return u;
    };
    const sentences = text.match(/[^.!?]+[.!?]*/g) || [text];
    sentences.forEach(s => { const t = s.trim(); if (t) window.speechSynthesis.speak(makeU(t)); });
  },

  _speakCurrent() {
    const { queue, qIdx } = this.s;
    if (qIdx < queue.length) this._speak(queue[qIdx].front);
  },

  _toggleAudio() {
    this.s.audioOn = !this.s.audioOn;
    localStorage.setItem('pgads_audio', this.s.audioOn);
    if (!this.s.audioOn) {
      window.speechSynthesis?.cancel();
      if (this._ttsAudio) { this._ttsAudio.pause(); this._ttsAudio = null; }
    }
    this._render('review');
  },

  _toggleMic() {
    this.s.isListening ? this._stopListening() : this._startListening();
  },

  _startListening() {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) { this._toast('Speech recognition not supported in this browser.', 'error'); return; }

    this._recognition = new SR();
    this._recognition.continuous = true;
    this._recognition.interimResults = true;
    this._recognition.lang = 'en-US';

    this._recognition.onstart = () => {
      this.s.isListening = true;
      const btn = document.getElementById('mic-btn');
      if (btn) { btn.textContent = '⏹'; btn.classList.add('listening'); btn.title = 'Stop recording'; }
    };

    this._recognition.onresult = (e) => {
      let finalText = this.s.userAnswer;
      let interim = '';
      for (let i = e.resultIndex; i < e.results.length; i++) {
        if (e.results[i].isFinal) {
          finalText += (finalText ? ' ' : '') + e.results[i][0].transcript;
        } else {
          interim += e.results[i][0].transcript;
        }
      }
      this.s.userAnswer = finalText;
      const ta = document.getElementById('user-answer');
      if (ta) ta.value = finalText + (interim ? ' ' + interim : '');
    };

    this._recognition.onerror = (e) => {
      if (e.error !== 'aborted') this._toast('Mic error: ' + e.error, 'error');
      this._stopListening();
    };

    this._recognition.onend = () => {
      this.s.isListening = false;
      const btn = document.getElementById('mic-btn');
      if (btn) { btn.textContent = '🎤'; btn.classList.remove('listening'); btn.title = 'Voice input'; }
    };

    this._recognition.start();
  },

  _stopListening() {
    if (this._recognition) {
      try { this._recognition.stop(); } catch {}
      this._recognition = null;
    }
    this.s.isListening = false;
  },

  // ── Toast ────────────────────────────────────────────────────────────────────
  _toast(msg, type = '') {
    let box = document.querySelector('.toast-container');
    if (!box) { box = document.createElement('div'); box.className = 'toast-container'; document.body.appendChild(box); }
    const t = document.createElement('div');
    t.className = `toast ${type}`;
    t.textContent = msg;
    box.appendChild(t);
    setTimeout(() => t.remove(), 3200);
  }
};

// ─── Helpers ──────────────────────────────────────────────────────────────────
function h(str) {
  return String(str || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// Tags are stored as: tagPath1 :: tagPath2 :: tagPath3
// Each tagPath uses :: for hierarchy: java::collections::list
// Separator between distinct tags is ' :: ' (with spaces).
// Legacy fallback: comma-separated.
function parseTags(str) {
  if (!str?.trim()) return [];
  if (str.includes(' :: ')) return str.split(' :: ').map(t => t.trim()).filter(Boolean);
  return str.split(',').map(t => t.trim()).filter(Boolean);
}

// Split a hierarchical tag path into segments: 'java::collections::list' → ['java','collections','list']
function tagSegments(path) {
  return path.split('::').map(s => s.trim()).filter(Boolean);
}

// Human-readable display name: 'java::collections::list' → 'java › collections › list'
function tagDisplayName(path) {
  return tagSegments(path).join(' › ');
}

function googleSvg() {
  return `<svg width="18" height="18" viewBox="0 0 18 18" xmlns="http://www.w3.org/2000/svg">
    <path d="M17.64 9.2c0-.637-.057-1.251-.164-1.84H9v3.481h4.844c-.209 1.125-.843 2.078-1.796 2.716v2.259h2.908c1.702-1.567 2.684-3.875 2.684-6.615z" fill="#4285F4"/>
    <path d="M9 18c2.43 0 4.467-.806 5.956-2.184l-2.908-2.259c-.806.54-1.837.86-3.048.86-2.344 0-4.328-1.584-5.036-3.711H.957v2.332C2.438 15.983 5.482 18 9 18z" fill="#34A853"/>
    <path d="M3.964 10.706c-.18-.54-.282-1.117-.282-1.706s.102-1.166.282-1.706V4.962H.957C.347 6.175 0 7.55 0 9s.348 2.825.957 4.038l3.007-2.332z" fill="#FBBC05"/>
    <path d="M9 3.58c1.321 0 2.508.454 3.44 1.345l2.582-2.58C13.463.891 11.426 0 9 0 5.482 0 2.438 2.017.957 4.958L3.964 7.29C4.672 5.163 6.656 3.58 9 3.58z" fill="#EA4335"/>
  </svg>`;
}

// ── Boot ──────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => App.init());
