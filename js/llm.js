// Azure OpenAI API wrapper
// Endpoint pattern: {endpoint}/openai/deployments/{deployment}/chat/completions?api-version={version}
const LLM = {
  _url() {
    const ep  = Config.azureEndpoint.replace(/\/$/, '');
    const dep = Config.azureDeployment;
    const ver = Config.azureApiVersion;
    return `${ep}/openai/deployments/${dep}/chat/completions?api-version=${ver}`;
  },

  _headers() {
    return {
      'Content-Type': 'application/json',
      'api-key': Config.azureApiKey
    };
  },

  async _call(messages, maxTokens = 512) {
    if (!Config.azureApiKey) {
      throw new Error('Azure OpenAI API key not set — add row azureApiKey / <key> to your Sheet\'s Settings tab, or enter it via the Settings page.');
    }
    const res = await fetch(this._url(), {
      method: 'POST',
      headers: this._headers(),
      body: JSON.stringify({ messages, max_completion_tokens: maxTokens })
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error?.message || `HTTP ${res.status}`);
    }

    const data = await res.json();
    return data.choices[0].message.content;
  },

  // Explain a flashcard in 2-3 sentences
  async explain(front, back) {
    try {
      return await this._call([{
        role: 'user',
        content: `Flashcard:\nQ: ${front}\nA: ${back}\n\nGive a concise explanation (2-3 sentences) to help me understand and remember this. Be practical and direct.`
      }], 400);
    } catch (e) {
      return `Error: ${e.message}`;
    }
  },

  // Validate user's answer against the question and correct answer
  async validate(question, correctAnswer, userAnswer) {
    try {
      return await this._call([{
        role: 'user',
        content: `You are evaluating a technical flashcard answer.\n\nQuestion: ${question}\n${correctAnswer ? `Correct answer: ${correctAnswer}\n` : ''}My answer: ${userAnswer}\n\nEvaluate concisely:\n1. Verdict: Correct / Partially correct / Incorrect\n2. What I got right (skip if nothing)\n3. What I missed or got wrong (skip if nothing)\n4. One key point to remember\n\nBe direct, 4-6 sentences max.`
      }], 450);
    } catch (e) {
      return 'Error: ' + e.message;
    }
  },

  // Text-to-speech via Azure OpenAI TTS — returns a blob URL for playback
  async tts(text, dep = Config.azureTtsDeployment, voice = Config.ttsVoice, epOverride = null, keyOverride = null) {
    const ep  = (epOverride || Config.azureTtsEndpoint || Config.azureEndpoint).replace(/\/$/, '');
    const ver = Config.azureTtsApiVersion;
    const key = keyOverride || Config.azureTtsApiKey || Config.azureApiKey;
    const res = await fetch(`${ep}/openai/deployments/${dep}/audio/speech?api-version=${ver}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` },
      body: JSON.stringify({ model: dep, input: text, voice, response_format: 'mp3' })
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error?.message || `TTS HTTP ${res.status}`);
    }
    const blob = await res.blob();
    return URL.createObjectURL(blob);
  },

  // Generate N cards about a topic — returns array of {front, back, tags}
  async generateCards(topic, count = 5) {
    const text = await this._call([{
      role: 'user',
      content: `Generate ${count} technical interview flashcards about: "${topic}"

Return ONLY a valid JSON array — no explanation, no markdown fences:
[{"front":"question","back":"answer","tags":"tag1,tag2"},...]

Questions should be concise. Answers should be complete but not verbose. Tags comma-separated.`
    }], 1800);

    const match = text.match(/\[[\s\S]*\]/);
    if (!match) throw new Error('Could not parse AI response as JSON array.');
    return JSON.parse(match[0]);
  }
};
