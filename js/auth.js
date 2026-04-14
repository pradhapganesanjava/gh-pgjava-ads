// Google Identity Services — OAuth 2.0 token flow
const Auth = {
  token: null,
  user:  null,

  // Attempt to restore session from sessionStorage
  async init() {
    if (!Config.googleClientId) return false;
    await Sheets.init();

    const raw = sessionStorage.getItem('pgads_tok');
    if (!raw) return false;
    try {
      const { token, expires } = JSON.parse(raw);
      if (Date.now() < expires) {
        this.token = token;
        this.user  = JSON.parse(sessionStorage.getItem('pgads_usr') || 'null');
        Sheets.setToken(token);
        Media.setToken(token);
        return true;
      }
    } catch {}
    return false;
  },

  signIn() {
    return new Promise((resolve, reject) => {
      if (!Config.googleClientId) {
        reject(new Error('Google Client ID not configured — open Settings first.'));
        return;
      }

      const client = google.accounts.oauth2.initTokenClient({
        client_id: Config.googleClientId,
        scope: [
          'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive.metadata.readonly',
          'https://www.googleapis.com/auth/drive.file',
          'https://www.googleapis.com/auth/userinfo.profile',
          'https://www.googleapis.com/auth/userinfo.email'
        ].join(' '),
        callback: async (res) => {
          if (res.error) { reject(new Error(res.error)); return; }

          this.token = res.access_token;
          Sheets.setToken(this.token);
          Media.setToken(this.token);

          const expires = Date.now() + res.expires_in * 1000;
          sessionStorage.setItem('pgads_tok', JSON.stringify({ token: this.token, expires }));

          // Fetch profile
          try {
            const r = await fetch('https://www.googleapis.com/oauth2/v1/userinfo', {
              headers: { Authorization: `Bearer ${this.token}` }
            });
            this.user = await r.json();
            sessionStorage.setItem('pgads_usr', JSON.stringify(this.user));
          } catch {}

          resolve(this.user);
        }
      });

      client.requestAccessToken({ prompt: '' });
    });
  },

  signOut() {
    if (this.token) {
      try { google.accounts.oauth2.revoke(this.token); } catch {}
    }
    this.token = null;
    this.user  = null;
    sessionStorage.removeItem('pgads_tok');
    sessionStorage.removeItem('pgads_usr');
  },

  isSignedIn() { return !!this.token; }
};
