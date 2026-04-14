// SM-2 Spaced Repetition Algorithm
const SM2 = {
  // User grades: 1=Again  2=Hard  3=Good  4=Easy
  // Maps to SM-2 quality scores (0-5 scale)
  QUALITY: { 1: 1, 2: 3, 3: 4, 4: 5 },

  defaultProgress() {
    return {
      easeFactor: 2.5,
      interval: 0,
      repetitions: 0,
      dueDate: today()
    };
  },

  // Returns updated progress after a review
  review(progress, grade) {
    const q = this.QUALITY[grade] || 3;
    let { easeFactor, interval, repetitions } = progress;

    if (q < 3) {
      // Failed — reset
      repetitions = 0;
      interval = 1;
    } else {
      if (repetitions === 0)      interval = 1;
      else if (repetitions === 1) interval = 6;
      else                        interval = Math.round(interval * easeFactor);

      easeFactor += 0.1 - (5 - q) * (0.08 + (5 - q) * 0.02);
      easeFactor = Math.max(1.3, Math.round(easeFactor * 100) / 100);
      repetitions++;
    }

    const due = new Date();
    due.setDate(due.getDate() + interval);

    return {
      easeFactor,
      interval,
      repetitions,
      dueDate: due.toISOString().split('T')[0],
      lastReview: today()
    };
  },

  // Preview next intervals for all 4 grades (for rating button labels)
  previewIntervals(progress) {
    return [1, 2, 3, 4].map(g => this.review(progress || this.defaultProgress(), g).interval);
  },

  formatInterval(days) {
    if (days <= 0)  return '<1d';
    if (days === 1) return '1d';
    if (days < 30)  return `${days}d`;
    if (days < 365) return `${Math.round(days / 30)}mo`;
    return `${Math.round(days / 365)}y`;
  },

  isDue(progress) {
    if (!progress || progress.repetitions === 0) return true;
    return progress.dueDate <= today();
  }
};

function today() {
  return new Date().toISOString().split('T')[0];
}
