// Detects another open tab of the same page in this browser. The pilot writer model allows one writer at a
// time; a second tab of the same browser is the one case the app itself can see, so it refuses to write
// from it. Other devices, Excel and the old script buttons are outside what any page can detect.
const randomId = () => `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;

export function createTabGuard({ channelName = 'tpaha-books', id = randomId(), channelFactory, timeoutMs = 400, setTimeoutImpl = setTimeout } = {}) {
  const factory = channelFactory || (name => (typeof BroadcastChannel === 'function' ? new BroadcastChannel(name) : null));
  const channel = factory(channelName);
  const state = { id, primary: true, others: 0, started: false, settled: false, listeners: [] };
  const post = msg => { try { channel && channel.postMessage({ ...msg, id }); } catch { /* ignore */ } };
  const onMessage = ev => {
    const m = ev && ev.data;
    if (!m || m.id === id) return;
    if (m.type === 'hello') {
      // A newcomer. If we are already primary we answer; if both are still starting, the lower id wins.
      if (state.settled) { if (state.primary) post({ type: 'present' }); return; }
      if (m.id < id) { state.primary = false; state.others += 1; } else post({ type: 'present' });
    } else if (m.type === 'present') {
      if (!state.settled) { state.primary = false; state.others += 1; }
      else if (state.primary) { for (const cb of state.listeners) cb(m); }
    }
  };
  if (channel) channel.onmessage = onMessage;
  return {
    id,
    get primary() { return state.primary; },
    get others() { return state.others; },
    /** Resolves { primary, others } after the discovery window. Without BroadcastChannel it resolves primary. */
    start() {
      if (state.started) return Promise.resolve({ primary: state.primary, others: state.others });
      state.started = true;
      if (!channel) { state.settled = true; return Promise.resolve({ primary: true, others: 0, unsupported: true }); }
      post({ type: 'hello' });
      return new Promise(resolve => setTimeoutImpl(() => { state.settled = true; resolve({ primary: state.primary, others: state.others }); }, timeoutMs));
    },
    /** Called on a primary tab when another tab later claims to be present (should not happen; informational). */
    onOtherTab(cb) { state.listeners.push(cb); },
    stop() { try { channel && channel.close(); } catch { /* ignore */ } },
    _deliver: onMessage,   // for tests with a fake channel
  };
}
