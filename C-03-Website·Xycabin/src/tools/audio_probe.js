/* Run with Node.js: isolated DOM/AudioContext doubles, no browser or playback. */
const fs = require("fs"), path = require("path"), vm = require("vm"), assert = require("assert");
const root = path.resolve(__dirname, "../.."), html = fs.readFileSync(process.argv[2] || path.join(root, "AA-Program/XyCabin.html"), "utf8");
const data = JSON.parse(html.match(/<script id="site-data" type="application\/json">([\s\S]*?)<\/script>/)[1]);
const source = fs.readFileSync(path.join(root, "src/site_audio.js"), "utf8");
let checks = 0;
const defaultFx = Math.round(data.sound.gain / data.sound.max_gain * 100), defaultBg = Math.round(data.background.volume * 100);
function check(value, name) { assert(value, name); checks++; }
class Element {
  constructor() { this.events = {}; this.children = []; this.style = {}; this.attrs = {}; this.hidden = true; }
  appendChild(child) { this.children.push(child); child.parent = this; }
  setAttribute(key, value) { this.attrs[key] = value; }
  addEventListener(key, fn) { this.events[key] = fn; }
  fire(key) { if (this.events[key]) this.events[key].call(this, {}); }
  contains(target) { return target === this || this.children.some(c => c.contains(target)); }
  focus() { this.focused = true; }
}
function harness(options = {}) {
  const elements = {}, nodes = [], contexts = [], timers = new Set();
  const input = JSON.parse(JSON.stringify(data));
  input.background.autoplay = options.automatic === true;
  if (options.fallbackDisabled) input.background.play_on_first_interaction = false;
  if (options.recordingDefault) input.background.default_scene = "tranquility";
  const get = id => elements[id] || (elements[id] = new Element());
  const doc = {getElementById: get, createElement: () => new Element(), addEventListener: () => {}};
  // Dynamic scene controls obtain IDs when added to the scene container.
  get("audio-scenes").appendChild = function (label) { this.children.push(label); label.parent = this; elements[label.children[0].id] = label.children[0]; };
  const param = () => ({value: 1, setValueAtTime(v) { this.value = v; }, linearRampToValueAtTime(v) { this.value = v; }, exponentialRampToValueAtTime(v) { this.value = v; }, cancelScheduledValues() {}, setTargetAtTime(v) { this.value = v; }});
  const make = kind => { const n = {kind, gain: param(), frequency: param(), connect() {}, disconnect() { this.disconnected = true; }, start(when,offset) { this.started = true; this.offset = offset; }, stop() { this.stopped = true; }}; nodes.push(n); return n; };
  let release, releaseDecode, decodeCount = 0, resumeCount = 0;
  class Audio {
    constructor() { if (options.constructorFails) throw Error("unavailable"); this.state = options.pending || options.refusesResume ? "suspended" : "running"; this.sampleRate = 1000; this.currentTime = 0; this.destination = {}; contexts.push(this); }
    createGain() { return make("gain"); }
    createOscillator() { return make("oscillator"); }
    createBufferSource() { return make("buffer"); }
    decodeAudioData(bytes, resolve, reject) { decodeCount++; if (options.decodeFails) { reject(Error("bad MP3")); return; } const done = () => resolve({duration: 51}); if (options.decodePending) releaseDecode = done; else done(); }
    createBiquadFilter() { return make("filter"); }
    close() { this.state = "closed"; return Promise.resolve(); }
    createDynamicsCompressor() { if (options.graphFails) throw Error("graph failure"); return {threshold: param(), knee: param(), ratio: param(), connect() {}}; }
    createBuffer(channels, length) { const samples = Array.from({length: channels}, () => new Float32Array(length)); return {getChannelData: i => samples[i]}; }
    resume() { resumeCount++; if (options.refusesResume) return Promise.resolve(); if (options.pending) return new Promise(resolve => { release = () => { this.state = "running"; resolve(); }; }); this.state = "running"; return Promise.resolve(); }
  }
  const store = Object.assign({}, options.saved || {}), events = {};
  const win = {AudioContext: options.noAudio ? undefined : Audio, performance: {now: () => Date.now()},
    atob: b => Buffer.from(b, "base64").toString("binary"),
    localStorage: {getItem(k) { if (options.storageBlocked) throw Error("blocked"); return store[k] || null; }, setItem(k,v) { if (options.storageBlocked) throw Error("blocked"); store[k] = v; }},
    setInterval(fn) { const token = {fn}; timers.add(token); return token; }, clearInterval(t) { timers.delete(t); }, addEventListener(k,fn) { events[k] = fn; }};
  const sandbox = {Set, Promise, Math, Number}; vm.createContext(sandbox); vm.runInContext(source, sandbox);
  const api = sandbox.createCabinAudio(input, doc, win);
  return {get, nodes, contexts, timers, store, events, api, release: () => release(), releaseDecode: () => releaseDecode(), decodes: () => decodeCount, resumes: () => resumeCount};
}
async function flush() { for (let i = 0; i < 12; i++) await Promise.resolve(); }
(async () => {
  let h = harness(); check(h.contexts.length === 0, "no audio or autoplay at page load");
  check(h.get("effects-volume").value === defaultFx && h.get("background-volume").value === defaultBg, "raised initial effects and independent volume defaults");
  h.get("sound-toggle").fire("click"); check(h.get("audio-panel").hidden === false && h.get("sound-toggle").attrs["aria-expanded"] === "true", "volume button opens panel");
  for (const s of data.background.scenes) {
    h.get("audio-scene-" + s.id).fire("change"); h.get("music-play").fire("click"); await flush();
    check(h.get("music-play").textContent === data.text.music_pause && h.get("audio-status").textContent.includes(s.name), "plays selected scene " + s.id);
    check(h.timers.size === (s.id === "afternoon" ? 1 : 0), "one synth scheduler or looped MP3"); h.get("music-play").fire("click");
    check(h.timers.size === 0 && h.nodes.filter(n => n.started).every(n => n.stopped), "pause releases all scene sources");
  }
  h.get("effects-enabled").checked = true; h.get("effects-enabled").fire("change");
  h.get("effects-volume").value = "100"; h.get("effects-volume").fire("input");
  h.get("background-volume").value = "0"; h.get("background-volume").fire("input");
  h.api.click(); await flush(); check(h.nodes[1].gain.value === data.sound.max_gain && h.nodes[2].gain.value === 0, "slider reaches higher effects ceiling without changing background volume");
  const count = h.nodes.filter(n => n.kind === "oscillator").length; h.api.click(); await flush();
  check(h.nodes.filter(n => n.kind === "oscillator").length === count, "continuous clicks rate limited");
  h.get("audio-mute").fire("click"); check(h.nodes[0].gain.value === 0, "all sound muted immediately through master gain");
  check(h.nodes.filter(n => n.kind === "oscillator" && n.started).every(n => n.stopped), "mute stops pending effects");
  let remembered = harness({saved: h.store}); check(remembered.get("effects-volume").value === 100 && remembered.get("background-volume").value === 0 && remembered.get("audio-mute").attrs["aria-pressed"] === "true", "valid volume and mute preferences restored");
  check(remembered.contexts.length === 0, "remembered settings never autoplay");
  h = harness({storageBlocked: true}); h.get("music-play").fire("click"); await flush(); check(h.get("music-play").attrs["aria-pressed"] === "true", "blocked storage does not prevent playback"); h.api.stop();
  h = harness({saved: {"xy-home-audio": '{"effects":"bad","background":-1,"scene":"bad","muted":0}', "xy-home-sound": "off"}});
  check(h.get("effects-volume").value === defaultFx && h.get("background-volume").value === defaultBg && !h.get("effects-enabled").checked, "invalid settings fall back and old effects preference retained");
  h = harness({saved: {"xy-home-audio": "malformed JSON", "xy-home-sound": "off"}});
  check(!h.get("effects-enabled").checked, "malformed new settings preserve old explicit effects mute");
  const fault = {graphFails: true}; h = harness(fault); h.get("music-play").fire("click"); await flush();
  check(h.contexts[0].state === "closed" && h.get("audio-status").textContent === data.text.audio_failure, "partial audio graph failure releases broken context");
  h.get("background-volume").value = "75"; h.get("background-volume").fire("input"); fault.graphFails = false;
  h.get("music-play").fire("click"); await flush();
  check(h.contexts.length === 2 && h.get("music-play").attrs["aria-pressed"] === "true", "graph can retry successfully and volume controls remain usable"); h.api.stop();
  h = harness({pending: true}); h.get("music-play").fire("click"); h.get("music-play").fire("click"); h.release(); await flush();
  check(h.timers.size === 0 && !h.nodes.some(n => n.started), "pause invalidates pending resume");
  h = harness({pending: true}); h.get("effects-enabled").checked = true; h.get("effects-enabled").fire("change"); h.get("effects-volume").value = "50"; h.get("effects-volume").fire("input"); h.api.click(); h.get("audio-mute").fire("click"); h.release(); await flush(); check(!h.nodes.some(n => n.started), "mute invalidates delayed effects");
  h = harness(); h.get("music-play").fire("click"); await flush();
  for (const s of data.background.scenes) { h.get("audio-scene-" + s.id).fire("change"); await flush(); check(h.timers.size === (s.id === "afternoon" ? 1 : 0) && h.get("audio-status").textContent.includes(s.name), "live scene switch keeps one active scene " + s.id); }
  h.events.pagehide(); check(h.timers.size === 0, "page departure releases background sound");
  for (const options of [{noAudio: true}, {constructorFails: true}, {refusesResume: true}]) {
    h = harness(options); h.get("music-play").fire("click"); await flush();
    check(h.timers.size === 0 && [data.text.sound_unavailable, data.text.audio_failure].includes(h.get("audio-status").textContent), "unavailable audio fails without hanging or playing");
  }
  check(data.background.scenes.map(s => s.id).join(",") === "afternoon,tranquility", "only two supported music choices");
  check(data.background.scenes[1].audio_base64.length > 100000, "cropped recording embedded in HTML");
  h = harness({automatic: true}); await flush();
  check(h.get("music-play").attrs["aria-pressed"] === "true", "allowed page-load autoplay starts music");
  const started = h.nodes.filter(n => n.started).length; h.api.interaction({isTrusted:true}); await flush();
  check(h.nodes.filter(n => n.started).length === started, "normal click never restarts successful autoplay"); h.api.stop();
  h = harness({automatic: true, refusesResume: true}); await flush();
  check(h.get("audio-status").textContent === data.text.audio_paused && !h.nodes.some(n => n.started), "blocked autoplay stays silently paused");
  h.contexts[0].state = "running"; h.api.click(); await flush();
  check(!h.nodes.some(n => n.kind === "buffer" && n.started) && h.timers.size === 0, "unrelated effects click never retries blocked music");
  h.get("music-play").fire("click"); await flush();
  check(h.get("music-play").attrs["aria-pressed"] === "true", "manual playback works after autoplay denial"); h.api.stop();
  const blocked = {automatic:true, refusesResume:true, recordingDefault:true}; h = harness(blocked); await flush();
  h.api.interaction({isTrusted:false}); await flush(); check(h.resumes() === 0, "synthetic event cannot consume fallback");
  h.get("effects-enabled").checked = false; h.get("effects-enabled").fire("change"); blocked.refusesResume = false; blocked.pending = true;
  h.api.interaction({isTrusted:true}); h.api.click(); h.api.interaction({isTrusted:true});
  check(h.resumes() === 1 && h.contexts.length === 1, "trusted click resumes immediately once with effects disabled");
  h.release(); await flush(); check(h.nodes.filter(n => n.kind === "buffer" && n.started).length === 1 && h.decodes() === 1, "consecutive clicks create only one decoded music source"); h.api.stop();
  const shared = {automatic:true, refusesResume:true, recordingDefault:true}; h = harness(shared); await flush(); shared.refusesResume = false; shared.pending = true;
  h.api.interaction({isTrusted:true}); h.api.click(); check(h.resumes() === 1, "effects and music share pending resume request"); h.release(); await flush(); h.api.stop();
  h = harness({automatic:true, refusesResume:true, recordingDefault:true}); await flush(); h.api.interaction({isTrusted:true}); await flush();
  check(h.resumes() === 1 && h.get("audio-status").textContent === data.text.audio_paused, "failed fallback is quiet");
  h.api.interaction({isTrusted:true}); await flush(); check(h.resumes() === 1, "failed fallback is never retried by normal clicks");
  h.contexts[0].state = "running"; h.get("music-play").fire("click"); await flush(); check(h.nodes.some(n => n.kind === "buffer" && n.started), "manual playback can retry after failed fallback"); h.api.stop();
  for (const cancel of ["mute", "zero", "scene", "pause", "pagehide"]) {
    const opt = {automatic:true, refusesResume:true, recordingDefault:true}; h = harness(opt); await flush();
    if (cancel === "mute") h.get("audio-mute").fire("click");
    if (cancel === "zero") { h.get("background-volume").value = "0"; h.get("background-volume").fire("input"); }
    if (cancel === "scene") h.get("audio-scene-afternoon").fire("change");
    if (cancel === "pause") h.api.stop();
    if (cancel === "pagehide") h.events.pagehide();
    opt.refusesResume = false; h.api.interaction({isTrusted:true}); await flush();
    check(!h.nodes.some(n => n.started) && h.resumes() === 0, cancel + " disarms pending first-click fallback");
  }
  const late = {automatic:true, refusesResume:true, recordingDefault:true, decodePending:true}; h = harness(late); await flush(); late.refusesResume = false;
  h.api.interaction({isTrusted:true}); await flush(); h.get("music-play").fire("click"); h.releaseDecode(); await flush();
  check(!h.nodes.some(n => n.started), "pause during fallback decoding prevents late playback");
  h = harness({automatic:true, refusesResume:true, fallbackDisabled:true}); await flush(); h.api.interaction({isTrusted:true}); await flush(); check(h.resumes() === 0, "disabled config suppresses first-click fallback");
  h = harness(); h.api.interaction({isTrusted:true}); await flush(); check(h.contexts.length === 0, "autoplay off suppresses first-click fallback");
  const bad = {automatic:true, refusesResume:true, recordingDefault:true, decodeFails:true}; h = harness(bad); await flush(); bad.refusesResume = false; h.api.interaction({isTrusted:true}); await flush();
  check(!h.nodes.some(n => n.started) && h.get("audio-status").textContent === data.text.audio_paused, "fallback decode failure is quiet and canceled");
  for (const saved of [{effects:100, background:35, muted:false}, {effects:50, background:0, muted:true}]) {
    h = harness({saved:{"xy-home-audio":JSON.stringify(saved)}});
    check(h.get("effects-volume").value === saved.effects/2 && h.get("background-volume").value === saved.background/2, "both legacy volumes migrate at unchanged gain");
    check(h.get("background-value").textContent === saved.background/2 + "%", "fractional percent displayed accurately");
    let again = harness({saved:h.store}); check(again.get("effects-volume").value === saved.effects/2, "migration never repeats");
  }
  for (const saved of [{version:2, effects:0, background:50, muted:true}, {version:2,effects:40,background:0,muted:false}]) {
    h = harness({automatic:true,saved:{"xy-home-audio":JSON.stringify(saved)}}); await flush();
    check(h.contexts.length === 0, "saved mute or zero background prevents autoplay");
  }
  h = harness(); h.get("audio-scene-tranquility").fire("change"); h.get("background-volume").value = "100"; h.get("background-volume").fire("input"); h.get("music-play").fire("click"); await flush();
  check(h.nodes[2].gain.value === data.background.max_gain, "background slider reaches doubled ceiling");
  h.contexts[0].currentTime = 12; h.get("music-play").fire("click"); h.get("music-play").fire("click"); await flush();
  check(h.decodes() === 1, "MP3 decoded once across pause/resume"); check(h.nodes.filter(n=>n.kind==="buffer").at(-1).offset===12,"pause resumes MP3 from saved offset"); h.api.stop();
  h = harness({automatic:true, decodePending:true, recordingDefault:true}); await flush(); h.api.stop(); h.releaseDecode(); await flush();
  check(!h.nodes.some(n=>n.started), "late decode cannot restart canceled autoplay");
  h = harness({decodeFails:true}); h.get("audio-scene-tranquility").fire("change"); h.get("music-play").fire("click"); await flush();
  check(h.get("audio-status").textContent === data.text.audio_failure, "manual decoder failure visible without disabling controls");
  h.get("audio-scene-afternoon").fire("change"); h.get("music-play").fire("click"); await flush(); check(h.timers.size===1, "synth still usable after decoder failure"); h.api.stop();
  console.log("AudioProbe: " + checks + " PASS, 0 FAIL (isolated logic, no browser or listening claim)");
})().catch(error => { console.error(error); process.exitCode = 1; });
