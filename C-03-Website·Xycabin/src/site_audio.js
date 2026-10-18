/* Embedded by site_core.py; embedded MP3 data; no network or adjacent runtime files. */
function createCabinAudio(data, doc, win) {
  "use strict";
  var T = data.text, cfg = data.sound, bg = data.background;
  var el = function (name) { return doc.getElementById(name); };
  var AudioClass = win.AudioContext || win.webkitAudioContext;
  var ctx = null, master = null, effects = null, ambience = null;
  var enabled = cfg.enabled, fxVolume = cfg.gain / cfg.max_gain * 100;
  var bgVolume = bg.volume * 100, scene = bg.default_scene, muted = false;
  var playing = false, failure = false, fxEpoch = 0, musicEpoch = 0, lastClick = -Infinity;
  var pulses = new Set(), group = null, pending = false, offset = 0, decoded = {};
  var fallbackEligible = false, resumeRequest = null;
  var button = el("sound-toggle"), panel = el("audio-panel");
  function validVolume(value) { return typeof value === "number" && Number.isFinite(value) && value >= 0 && value <= 100; }
  try {
    var saved = JSON.parse(win.localStorage.getItem("xy-home-audio") || "null");
    if (saved && typeof saved === "object") {
      var divisor = saved.version === 2 ? 1 : 2;
      if (validVolume(saved.effects)) fxVolume = saved.effects / divisor;
      if (validVolume(saved.background)) bgVolume = saved.background / divisor;
      if (bg.scenes.some(function (s) { return s.id === saved.scene; })) scene = saved.scene;
      if (typeof saved.muted === "boolean") muted = saved.muted;
    }
  } catch (error) { /* Storage is optional, including on file:// pages. */ }
  // Malformed new preferences must not discard an existing explicit mute.
  try {
    var legacy = win.localStorage.getItem("xy-home-sound");
    if (legacy === "on" || legacy === "off") enabled = legacy === "on";
  } catch (error) {}
  function remember() {
    try {
      win.localStorage.setItem("xy-home-sound", enabled ? "on" : "off");
      win.localStorage.setItem("xy-home-audio", JSON.stringify({version: 2, effects: fxVolume, background: bgVolume, scene: scene, muted: muted}));
    } catch (error) {}
  }
  function smooth(param, value) {
    param.cancelScheduledValues(ctx.currentTime);
    param.setTargetAtTime(value, ctx.currentTime, 0.015);
  }
  function levels() {
    if (!ctx || ctx.state === "closed") return;
    smooth(master.gain, muted ? 0 : 1);
    smooth(effects.gain, enabled ? cfg.max_gain * fxVolume / 100 : 0);
    smooth(ambience.gain, bg.max_gain * bgVolume / 100);
  }
  function state() {
    el("effects-enabled").checked = enabled;
    el("effects-volume").value = fxVolume;
    el("background-volume").value = bgVolume;
    el("effects-value").textContent = Number(fxVolume.toFixed(2)) + "%";
    el("background-value").textContent = Number(bgVolume.toFixed(2)) + "%";
    el("music-play").textContent = playing || pending ? T.music_pause : T.music_play;
    el("music-play").setAttribute("aria-pressed", String(playing || pending));
    el("audio-mute").textContent = muted ? T.audio_unmute : T.audio_mute;
    el("audio-mute").setAttribute("aria-pressed", String(muted));
    el("sound-waves").style.display = muted ? "none" : "";
    el("sound-off").style.display = muted ? "" : "none";
    button.title = T.audio_title + (muted ? " · " + T.audio_muted : "");
    var selected = bg.scenes.find(function (s) { return s.id === scene; });
    el("audio-status").textContent = !AudioClass ? T.sound_unavailable : failure ? T.audio_failure :
      pending ? T.audio_starting : playing ? (muted ? T.audio_muted + " · " : "") + selected.name + " · " + T.audio_playing : T.audio_paused;
    bg.scenes.forEach(function (s) { el("audio-scene-" + s.id).checked = scene === s.id; });
    ["music-play", "audio-mute", "effects-enabled", "effects-volume", "background-volume"].forEach(function (name) { el(name).disabled = !AudioClass; });
  }
  function ready(automatic) {
    if (!ctx || ctx.state === "closed") {
      try {
        ctx = new AudioClass();
        master = ctx.createGain(); effects = ctx.createGain(); ambience = ctx.createGain();
        var limiter = ctx.createDynamicsCompressor();
        limiter.threshold.value = -10; limiter.knee.value = 14; limiter.ratio.value = 4;
        master.connect(limiter); limiter.connect(ctx.destination);
        effects.connect(master); ambience.connect(master); levels();
      } catch (error) {
        var broken = ctx; ctx = master = effects = ambience = null;
        try { if (broken) Promise.resolve(broken.close()).catch(function () {}); } catch (closeError) {}
        throw error;
      }
    }
    // resume() must be invoked inside the trusted event, before asynchronous decoding.
    if (automatic || ctx.state === "running") return Promise.resolve();
    if (resumeRequest) return resumeRequest;
    var request = Promise.resolve(ctx.resume()); resumeRequest = request;
    function complete() { if (resumeRequest === request) resumeRequest = null; }
    request.then(complete, complete); return request;
  }
  function stopEffects() {
    fxEpoch++;
    pulses.forEach(function (p) { try { p.osc.stop(); } catch (error) {} p.osc.disconnect(); p.gain.disconnect(); });
    pulses.clear();
  }
  function failed() { failure = true; state(); }
  function click() {
    if (!AudioClass || !enabled || muted || fxVolume === 0) return;
    var stamp = win.performance.now(); if (stamp - lastClick < 100) return;
    lastClick = stamp; var epoch = fxEpoch;
    try {
      Promise.resolve(ready()).then(function () {
        if (epoch !== fxEpoch || !enabled || muted || fxVolume === 0 || ctx.state !== "running") return;
        var osc = ctx.createOscillator(), gain = ctx.createGain(), now = ctx.currentTime, duration = cfg.duration_ms / 1000;
        var pulse = {osc: osc, gain: gain};
        osc.type = cfg.waveform; osc.frequency.setValueAtTime(cfg.frequency_hz, now);
        gain.gain.setValueAtTime(0.0001, now); gain.gain.linearRampToValueAtTime(1, now + 0.004);
        gain.gain.exponentialRampToValueAtTime(0.0001, now + duration);
        osc.connect(gain); gain.connect(effects); pulses.add(pulse);
        osc.onended = function () { osc.disconnect(); gain.disconnect(); pulses.delete(pulse); };
        osc.start(now); osc.stop(now + duration); failure = false; state();
      }).catch(failed);
    } catch (error) { failed(); }
  }
  function dispose(g) {
    if (!g) return; win.clearInterval(g.timer);
    g.sources.forEach(function (source) { try { source.stop(); } catch (error) {} });
    g.nodes.forEach(function (n) { try { n.disconnect(); } catch (error) {} });
    g.sources.clear(); g.nodes.clear();
  }
  function stopMusic(reset, preserveFallback) {
    if (!preserveFallback) fallbackEligible = false;
    if (group && group.buffer) offset = (group.offset + ctx.currentTime - group.started) % group.buffer.duration;
    if (reset) offset = 0;
    musicEpoch++; pending = false; playing = false; dispose(group); group = null; state();
  }
  function soundscape(buffer) {
    var g = {nodes: new Set(), sources: new Set(), timer: null};
    function own(n) { g.nodes.add(n); return n; }
    var bus = own(ctx.createGain()); bus.gain.setValueAtTime(0, ctx.currentTime);
    bus.gain.linearRampToValueAtTime(1, ctx.currentTime + 0.18); bus.connect(ambience);
    function tone(frequency, when, length, amount, wave) {
      var osc = own(ctx.createOscillator()), envelope = own(ctx.createGain());
      osc.type = wave || "sine"; osc.frequency.setValueAtTime(frequency, when);
      envelope.gain.setValueAtTime(0.0001, when);
      envelope.gain.exponentialRampToValueAtTime(amount, when + 0.025);
      envelope.gain.exponentialRampToValueAtTime(0.0001, when + length);
      osc.connect(envelope); envelope.connect(bus); g.sources.add(osc);
      osc.onended = function () { osc.disconnect(); envelope.disconnect(); g.sources.delete(osc); g.nodes.delete(osc); g.nodes.delete(envelope); };
      osc.start(when); osc.stop(when + length + 0.03);
    }
    function hz(midi) { return 440 * Math.pow(2, (midi - 69) / 12); }
    var next = ctx.currentTime + 0.05, beat = 0, horizon = 4;
    var chords = [[53,57,60,65],[50,53,57,62],[46,53,58,62],[48,55,60,64]];
    var melody = [69,72,74,72,69,67,65,67,69,65,62,65,67,69,72,69];
    function schedule() {
      if (ctx.state !== "running") return;
      if (next < ctx.currentTime) next = ctx.currentTime + 0.04;
      while (next < ctx.currentTime + horizon) {
        var chord = chords[Math.floor(beat / 8) % chords.length];
        tone(hz(chord[beat % 4]), next, 2.7, 0.29, "triangle");
        if (beat % 2 === 0) tone(hz(melody[Math.floor(beat / 2) % melody.length]), next + 0.1, 1.8, 0.2);
        if (beat % 8 === 0) tone(hz(chord[0] - 12), next, 4.5, 0.22);
        next += 60 / 68; beat++;
      }
    }
    try {
      if (buffer) {
        var source = own(ctx.createBufferSource()); source.buffer = buffer; source.loop = true;
        source.connect(bus); g.sources.add(source); g.buffer = buffer;
        g.offset = offset % buffer.duration; g.started = ctx.currentTime;
        source.start(0, g.offset); return g;
      }
      schedule(); g.timer = win.setInterval(function () { try { schedule(); } catch (error) { stopMusic(); failed(); } }, 300);
      return g;
    } catch (error) { dispose(g); throw error; }
  }
  function recording(selected) {
    if (!selected.audio_base64) return Promise.resolve(null);
    if (!decoded[selected.id]) {
      var binary = win.atob(selected.audio_base64), bytes = new Uint8Array(binary.length);
      for (var i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      decoded[selected.id] = new Promise(function (resolve, reject) {
        ctx.decodeAudioData(bytes.buffer, resolve, reject);
      }).catch(function (error) { delete decoded[selected.id]; throw error; });
    }
    return decoded[selected.id];
  }
  function startMusic(unmute, automatic, fromInteraction) {
    fallbackEligible = false;
    var epoch = ++musicEpoch; pending = true; failure = false; state();
    function reject(error) {
      if (epoch !== musicEpoch) return;
      stopMusic(false, true);
      fallbackEligible = !!(automatic && !error && ctx && ctx.state === "suspended" &&
        bg.autoplay && bg.play_on_first_interaction && !muted && bgVolume > 0);
      if (!automatic && !fromInteraction) failed();
    }
    try {
      Promise.resolve(ready(automatic)).then(function () {
        if (epoch !== musicEpoch || !pending) return;
        if (ctx.state !== "running") { reject(); return; }
        return recording(bg.scenes.find(function (s) { return s.id === scene; })).then(function (buffer) {
          if (epoch !== musicEpoch || !pending) return;
          if (ctx.state !== "running" || (automatic && (muted || bgVolume === 0))) { reject(); return; }
          var old = group; group = soundscape(buffer); dispose(old);
          pending = false; playing = true; if (unmute) muted = false;
          levels(); remember(); state();
        });
      }).catch(reject);
    } catch (error) { reject(error); }
  }
  bg.scenes.forEach(function (s) {
    var label = doc.createElement("label"), radio = doc.createElement("input"), text = doc.createElement("span"), name = doc.createElement("strong"), note = doc.createElement("small");
    label.className = "audio-scene"; radio.type = "radio"; radio.name = "audio-scene"; radio.id = "audio-scene-" + s.id; radio.value = s.id;
    name.textContent = s.name; note.textContent = s.description;
    text.appendChild(name); text.appendChild(note); label.appendChild(radio); label.appendChild(text); el("audio-scenes").appendChild(label);
    radio.addEventListener("change", function () { var resume = playing || pending; stopMusic(true); scene = s.id; remember(); if (resume) startMusic(false, false); else state(); });
  });
  function closePanel(restore) { panel.hidden = true; button.setAttribute("aria-expanded", "false"); if (restore) button.focus(); }
  button.setAttribute("aria-label", T.audio_title); button.setAttribute("aria-controls", "audio-panel"); button.setAttribute("aria-expanded", "false");
  button.addEventListener("click", function () { panel.hidden = !panel.hidden; button.setAttribute("aria-expanded", String(!panel.hidden)); });
  el("audio-close").addEventListener("click", function () { closePanel(true); });
  doc.addEventListener("click", function (e) { if (!panel.hidden && !panel.contains(e.target) && !button.contains(e.target)) closePanel(false); });
  doc.addEventListener("keydown", function (e) { if (e.key === "Escape" && !panel.hidden) { closePanel(true); e.preventDefault(); } });
  el("music-play").addEventListener("click", function () { if (playing || pending) stopMusic(); else startMusic(true, false); });
  el("audio-mute").addEventListener("click", function () { fallbackEligible = false; muted = !muted; if (muted) { stopEffects(); if (pending) stopMusic(); } levels(); remember(); state(); });
  el("effects-enabled").addEventListener("change", function () { enabled = this.checked; if (!enabled) stopEffects(); levels(); remember(); state(); });
  ["effects", "background"].forEach(function (kind) {
    el(kind + "-volume").addEventListener("input", function () {
      var value = Number(this.value); if (!validVolume(value)) return;
      if (kind === "effects") { fxVolume = value; if (value === 0) stopEffects(); } else { bgVolume = value; if (value === 0) { fallbackEligible = false; if (pending) stopMusic(); } }
      levels(); remember(); state();
    });
  });
  win.addEventListener("pagehide", function () { stopMusic(true); stopEffects(); });
  remember(); state();
  if (bg.autoplay && AudioClass && !muted && bgVolume > 0) startMusic(false, true);
  return {click: click, interaction: function (event) {
    if (!event || event.isTrusted !== true || !fallbackEligible || playing || pending || muted || bgVolume === 0) return;
    fallbackEligible = false; startMusic(false, false, true);
  }, stop: function () { stopMusic(); stopEffects(); }};
}
