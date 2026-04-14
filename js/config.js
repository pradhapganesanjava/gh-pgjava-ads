// Config — sheetId reading order: window.ENV_SHEET_ID (CI secret) → localStorage → ''
// Other values: localStorage cache → window.ENV_* → hardcoded default
// IMPORTANT: Azure API key is NOT hardcoded here. It is loaded from the
// "Settings" Google Sheet after sign-in and cached in localStorage.
const Config = {

  // ── In-memory backing for sensitive values (not in source code) ───
  // Seeded from localStorage cache on page load; overwritten from sheet after sign-in.
  _azureApiKey: localStorage.getItem('pgads_azk') || '',

  // ── Getters ───────────────────────────────────────────────────────
  get googleClientId() {
    return localStorage.getItem('pgads_gci')
      || window.ENV_GOOGLE_CLIENT_ID
      || '58187127990-a38oe06rl9jufk33bu1t82bctea8bevq.apps.googleusercontent.com';
  },
  get googleApiKey()   { return localStorage.getItem('pgads_gak') || ''; },
  get sheetId()        { return window.ENV_SHEET_ID || localStorage.getItem('pgads_sid') || '10Y9_L2_TnmNyWDUCPlZsrDF9nz04d0unCxHhaQxzeXA'; },

  // Google Drive folder ID for Anki media files
  get mediaDriveFolderId() { return localStorage.getItem('pgads_mdf') || ''; },

  // Azure OpenAI — endpoint/deployment/version are not secret, so defaults are fine
  get azureEndpoint()       { return localStorage.getItem('pgads_aze')   || ''; },
  get azureDeployment()     { return localStorage.getItem('pgads_azd')   || 'gpt-4o'; },
  get azureApiVersion()     { return localStorage.getItem('pgads_azv')   || '2024-12-01-preview'; },
  get azureTtsEndpoint()    { return localStorage.getItem('pgads_aztep')  || ''; },
  get azureTtsApiKey()      { return localStorage.getItem('pgads_aztak')  || ''; },
  get azureTtsDeployment()  { return localStorage.getItem('pgads_aztts')  || ''; },
  get azureTtsApiVersion()  { return localStorage.getItem('pgads_azttsv') || '2025-03-01-preview'; },
  get ttsVoice()            { return localStorage.getItem('pgads_ttsv')  || 'nova'; },

  // API key: in-memory only (loaded from sheet / cached in localStorage)
  // NEVER hardcoded — this getter returns '' if not yet loaded
  get azureApiKey() { return this._azureApiKey; },

  // UI theme
  get theme() { return localStorage.getItem('pgads_theme') || 'dark'; },

  // Speech voice (stored by voice name)
  get voiceName() { return localStorage.getItem('pgads_voice') || ''; },

  // ── Setters ───────────────────────────────────────────────────────
  set googleClientId(v) { localStorage.setItem('pgads_gci', v); },
  set googleApiKey(v)   { localStorage.setItem('pgads_gak', v); },
  set sheetId(v)        { localStorage.setItem('pgads_sid', v); },
  set mediaDriveFolderId(v) { localStorage.setItem('pgads_mdf', v); },

  set azureEndpoint(v)      { localStorage.setItem('pgads_aze',   v); },
  set azureDeployment(v)    { localStorage.setItem('pgads_azd',   v); },
  set azureApiVersion(v)    { localStorage.setItem('pgads_azv',   v); },
  set azureTtsEndpoint(v)    { localStorage.setItem('pgads_aztep',  v); },
  set azureTtsApiKey(v)      { localStorage.setItem('pgads_aztak',  v); },
  set azureTtsDeployment(v)  { localStorage.setItem('pgads_aztts',  v); },
  set azureTtsApiVersion(v)  { localStorage.setItem('pgads_azttsv', v); },
  set ttsVoice(v)            { localStorage.setItem('pgads_ttsv',   v); },

  // API key setter: update in-memory + cache locally (not in source, but in localStorage)
  set azureApiKey(v) {
    this._azureApiKey = v || '';
    if (v) localStorage.setItem('pgads_azk', v);
    else   localStorage.removeItem('pgads_azk');
  },

  set theme(v)     { localStorage.setItem('pgads_theme', v); },
  set voiceName(v) { localStorage.setItem('pgads_voice', v); },

  // Both Google Client ID and Sheet ID are required
  isConfigured() { return !!this.googleClientId && !!this.sheetId; }
};
