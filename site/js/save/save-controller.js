// Explicit Save workflow for one form or record.
// States: idle | unsaved | saving | saved | failed | conflict.
//
// One operation identity ("marker") per attempt sequence:
// - the first submit creates a pending operation { marker, payload } BEFORE the write; any failure keeps it;
// - check() is READ-ONLY: it asks `inspect` whether the pending operation landed and never writes;
// - submit() with a pending operation asks `resume` to finish it under the SAME marker (this may write the
//   finishing steps: formats, number check, sorted table, removal of an old copy) and writes the payload
//   again only if nothing landed;
// - the current draft is reported saved only when its content equals the payload that landed; a newer
//   draft stays visible and unsaved (it has its own identity later);
// - a submit while saving is ignored (no duplicate submissions);
// - edits typed during a save are kept; the state ends as `unsaved`, never `saved`.
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);

/**
 * Safe default operation id: the time as ISO 8601 UTC with nine fraction digits (millisecond + six random
 * digits). Two operations started in the same millisecond, or with a clock that does not move, still get
 * different ids (review finding T3). The ledger passes its own generator with the same format.
 */
export function defaultMarker() {
  const iso = new Date().toISOString();
  let digits = '';
  if (globalThis.crypto && typeof globalThis.crypto.getRandomValues === 'function') {
    const bytes = new Uint8Array(6);
    globalThis.crypto.getRandomValues(bytes);
    digits = Array.from(bytes, b => String(b % 10)).join('');
  } else {
    digits = String(Math.floor(Math.random() * 1e6)).padStart(6, '0');
  }
  return `${iso.slice(0, -1)}${digits}Z`;
}

export class SaveController {
  constructor({ perform, inspect, resume, onChange, makeMarker }) {
    if (typeof perform !== 'function') throw new Error('SaveController needs perform(draft, { marker })');
    this.perform = perform;
    this.inspect = inspect || (async () => ({ outcome: 'missing' }));
    this.resume = resume || (async () => null);
    this.onChange = onChange || (() => {});
    this.makeMarker = makeMarker || defaultMarker;
    this.state = 'idle';
    this.phase = null;            // 'checking' | 'resuming' | 'writing' while saving
    this.draft = null;
    this.generation = 0;
    this.savedGeneration = 0;
    this.pending = null;          // { marker, payload }
    this.error = null;
    this.conflict = null;
    this.result = null;
    this.needsVerification = false;
    this.incomplete = null;       // inspect() said: landed but finishing steps unconfirmed
  }

  get hasUnsavedEdits() {
    return (this.draft !== null && this.generation !== this.savedGeneration) || this.pending !== null;
  }

  _snapshot() {
    return { state: this.state, phase: this.phase, draft: this.draft, pending: this.pending, error: this.error, conflict: this.conflict, result: this.result, needsVerification: this.needsVerification, incomplete: this.incomplete, hasUnsavedEdits: this.hasUnsavedEdits };
  }
  _setState(state, phase = null) {
    this.state = state;
    this.phase = phase;
    this.onChange(this._snapshot());
  }

  /** Record what the member typed. Safe to call while a save is in flight. */
  edit(draft) {
    this.draft = draft;
    this.generation += 1;
    if (this.state !== 'saving') this._setState('unsaved');
    else this.onChange(this._snapshot());
  }

  /** A pending operation has landed. The draft counts as saved only if its content is what landed. */
  _landed(found) {
    const p = this.pending;
    this.pending = null;
    this.result = found;
    this.incomplete = null;
    const current = this.draft === null || same(this.draft, p.payload);
    if (current) { this.savedGeneration = this.generation; this._setState('saved'); }
    else this._setState('unsaved');
    return { ok: true, result: found, resolvedPending: true, currentDraftSaved: current };
  }

  _isConflict(e) { return Boolean(e && (e.name === 'ConflictError' || e.name === 'MarkerCollisionError')); }

  /** Save. Returns { ok, result, resolvedPending? } | { ok:false, error } | { ignored, reason }. */
  async submit() {
    if (this.state === 'saving') return { ignored: true, reason: 'saving' };
    if (this.draft === null && !this.pending) return { ignored: true, reason: 'nothing' };
    const gen = this.generation;
    const draft = this.draft;
    this.error = null;
    this.conflict = null;
    this.needsVerification = false;
    this.incomplete = null;
    try {
      let marker;
      if (this.pending) {
        this._setState('saving', 'resuming');
        const found = await this.resume(this.pending.payload, { marker: this.pending.marker });
        if (found) return this._landed(found);
        marker = this.pending.marker;                       // nothing landed: same identity, current content
        if (draft === null) { this.error = new Error('Nothing to save.'); this._setState('failed'); return { ok: false, error: this.error }; }
      } else {
        marker = this.makeMarker();
      }
      this.pending = { marker, payload: draft };
      this._setState('saving', 'writing');
      const result = await this.perform(draft, { marker });
      this.pending = null;
      this.result = result;
      if (this.generation === gen) { this.savedGeneration = gen; this._setState('saved'); }
      else this._setState('unsaved');
      return { ok: true, result };
    } catch (e) {
      if (this._isConflict(e)) { this.pending = null; this.conflict = e; this._setState('conflict'); }
      else { this.error = e; this.needsVerification = Boolean(e && e.ambiguous); this._setState('failed'); }
      return { ok: false, error: e };
    }
  }

  /**
   * Read-only: did the pending operation land? Never writes.
   * inspect() answers { outcome: 'missing' | 'landed' | 'incomplete' | 'conflict', result?, error?, message? }.
   */
  async check() {
    if (!this.pending) return { found: false, nothingPending: true };
    if (this.state === 'saving') return { ignored: true, reason: 'saving' };
    this.error = null;
    this.incomplete = null;
    this._setState('saving', 'checking');
    try {
      const r = (await this.inspect(this.pending.payload, { marker: this.pending.marker })) || { outcome: 'missing' };
      if (r.outcome === 'landed') { const l = this._landed(r.result); return { found: true, outcome: 'landed', ...l }; }
      if (r.outcome === 'incomplete') {
        this.incomplete = r;
        this.needsVerification = false;
        this.error = new Error(r.message || 'This is in the workbook, but its finishing steps are not confirmed. Press Save to finish; nothing is written twice.');
        this._setState('failed');
        return { found: true, outcome: 'incomplete', detail: r };
      }
      if (r.outcome === 'conflict') {
        this.pending = null;                                 // that operation can never complete as sent
        this.conflict = r.error || new Error(r.message || 'Someone else changed this record.');
        this._setState('conflict');
        return { found: false, outcome: 'conflict', conflict: this.conflict };
      }
      this.needsVerification = false;
      this.error = new Error('This is not in the workbook yet. Press Save to send it.');
      this._setState('failed');
      return { found: false, outcome: 'missing' };
    } catch (e) {
      this.error = e;
      this.needsVerification = true;
      this._setState('failed');
      return { found: false, error: e };
    }
  }

  retry() { return this.submit(); }

  /** After a reload or sign-in redirect: an attempt whose outcome was never confirmed. The current draft is left alone. */
  restorePending({ marker, payload }) {
    this.pending = { marker, payload };
    if (this.draft === null) { this.draft = payload; this.generation += 1; }
    this.needsVerification = true;
    this.error = new Error('An earlier save attempt has not been confirmed yet.');
    this._setState('failed');
  }

  reset({ discardPending = false } = {}) {
    this.draft = null;
    this.generation += 1;
    this.savedGeneration = this.generation;
    this.error = null;
    this.conflict = null;
    this.result = null;
    this.incomplete = null;
    if (discardPending) this.pending = null;
    if (this.pending) { this.needsVerification = true; this._setState('failed'); }
    else { this.needsVerification = false; this._setState('idle'); }
  }
}
