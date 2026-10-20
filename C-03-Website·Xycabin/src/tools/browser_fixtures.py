"""Produce isolated QA pages; never change the delivered website."""
from pathlib import Path
import copy
import json
import sys
sys.dont_write_bytecode=True
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT/'src'))
from site_core import load, render

script=r'''<script>
(function(){
var mode=__MODE__,real=window.AudioContext||window.webkitAudioContext,created=0,starts=0,last={};
function report(state){document.documentElement.setAttribute('data-qa-audio',JSON.stringify({created:created,starts:starts,state:state,last:last}))}
report('uninitialized');
if(mode==='storage'){Object.defineProperty(window,'localStorage',{get:function(){throw new DOMException('QA storage denied','SecurityError')}})}
else{try{localStorage.setItem('xy-home-theme',mode==='light'?'light':mode==='invalid'?'bad':'dark');localStorage.setItem('xy-home-sound','on')}catch(e){}}
if(mode==='unsupported'){window.AudioContext=undefined;window.webkitAudioContext=undefined}
else window.AudioContext=function(){created++;if(mode==='failure'){report('failed');throw new Error('QA audio unavailable')}
var ctx=new real(),createOsc=ctx.createOscillator.bind(ctx),createGain=ctx.createGain.bind(ctx);
ctx.createOscillator=function(){var osc=createOsc(),start=osc.start.bind(osc),stop=osc.stop.bind(osc),frequency=osc.frequency.setValueAtTime.bind(osc.frequency);
osc.frequency.setValueAtTime=function(v,t){last.frequency=v;return frequency(v,t)};osc.start=function(t){starts++;last.waveform=osc.type;report(ctx.state);return start(t)};osc.stop=function(t){if(t!==undefined)last.duration=Math.round((t-ctx.currentTime)*1000);report(ctx.state);return stop(t)};return osc};
ctx.createGain=function(){var gain=createGain(),ramp=gain.gain.linearRampToValueAtTime.bind(gain.gain);gain.gain.linearRampToValueAtTime=function(v,t){last.gain=v;return ramp(v,t)};return gain};return ctx};
if(mode==='manual'){Object.defineProperty(navigator,'clipboard',{value:{writeText:function(){return Promise.reject(new Error('QA copy denied'))}}});document.execCommand=function(){return false}}
else {if(navigator.clipboard&&navigator.clipboard.writeText){var write=navigator.clipboard.writeText.bind(navigator.clipboard);navigator.clipboard.writeText=function(value){return write(value).then(function(){document.documentElement.setAttribute('data-qa-copy',value)})}}
var copy=document.execCommand.bind(document);document.execCommand=function(command){var value=document.activeElement.value,ok=copy(command);if(command==='copy'&&ok)document.documentElement.setAttribute('data-qa-copy',value);return ok}}
window.print=function(){document.documentElement.setAttribute('data-qa-print',document.querySelector('#guide-tabs [aria-selected=true]').id)};
addEventListener('DOMContentLoaded',function(){var burst=document.createElement('button');burst.id='qa-burst';burst.textContent='QA 连续操作';burst.addEventListener('click',function(){for(var i=0;i<20;i++)document.getElementById('theme-toggle').click()});document.body.appendChild(burst)});
})();
</script>'''
config=vars(load(ROOT/'AA-Program/C3-XycabinSite_Configure.py'))
for mode in ('native','storage','invalid','light','failure','unsupported','manual','fallback'):
    selected=config
    if mode=='fallback':
        selected={key:copy.deepcopy(config[key]) for key in ('SITE','TEXT','PRODUCTS','WORKFLOW','FUTURE','THEME','SOUND','BACKGROUND','DOWNLOAD_GUIDE','DOWNLOADS','FEEDBACK')}
        selected['FEEDBACK']['emails']=['custom@example.org']
    html=render(selected,ROOT).replace('<script>try{',script.replace('__MODE__',json.dumps(mode))+'\n<script>try{',1)
    (ROOT/'build').mkdir(exist_ok=True)
    (ROOT/'build'/('_qa_'+mode+'.html')).write_text(html,encoding='utf-8')
print('Browser QA fixtures generated under build; delivered HTML unchanged')
