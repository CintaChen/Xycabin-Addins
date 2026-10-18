function createCabinDownloads(D, doc, win, selected, notify) {
 "use strict";
 var T=D.text, packages=JSON.parse(doc.getElementById("installer-assets").textContent).packages,
     main=doc.getElementById("download-main"), toggle=doc.getElementById("download-toggle"),
     menu=doc.getElementById("download-menu"), urls=[], busy=false,
     dialog=doc.getElementById("download-guide"), steps=doc.getElementById("guide-dialog-steps"),
     image=doc.getElementById("guide-dialog-image"), proceed=doc.getElementById("guide-dialog-continue"),
     guide=D.download_guide, pending=null, built=false;
 function close(){menu.hidden=true;toggle.setAttribute("aria-expanded","false")}
 function positionMenu(){
  if(menu.hidden)return;
  var viewport=win.visualViewport, width=viewport?viewport.width:doc.documentElement.clientWidth||win.innerWidth,
      left=(viewport?viewport.offsetLeft:0)+8, right=left+width-16;
  menu.style.maxWidth=Math.max(1,width-16)+"px";menu.style.transform="";
  var bounds=menu.getBoundingClientRect(),shift=Math.max(left-bounds.left,Math.min(0,right-bounds.right));
  menu.style.transform="translateX("+shift+"px)";
 }
 function product(){return selected()}
 function links(){return D.downloads[product()]||{}}
 function safeURL(value){try{var u=new win.URL(value);return u.protocol==="https:"&&!u.username&&!u.password&&!/[\s\x00-\x1f]/.test(value)?u.href:null}catch(e){return null}}
 function external(channel){var value=safeURL(links()[channel]);if(!value){notify(T.download_unconfigured);return false}var a=doc.createElement("a");a.href=value;a.target="_blank";a.rel="noopener noreferrer";doc.body.appendChild(a);a.click();a.remove();return true}
 function extract(){
  if(busy)return false;var pack=packages[product()];if(!pack){notify(T.download_unavailable);return false}
  busy=true;main.disabled=true;var url=null,a=null;
  try {
   var blocks=[], total=0, chunk=87380; // Multiple of four; each decoded string stays below 64 KiB.
   if(!Number.isSafeInteger(pack.bytes)||pack.bytes<=0||!/^\w+_Setup\.exe$/.test(pack.filename)||!/^[a-f0-9]{64}$/.test(pack.sha256))throw new Error("Invalid installer data");
   for(var i=0;i<pack.base64.length;i+=chunk){var binary=win.atob(pack.base64.slice(i,i+chunk)),part=new Uint8Array(binary.length);for(var j=0;j<binary.length;j++)part[j]=binary.charCodeAt(j);total+=part.length;blocks.push(part)}
   if(total!==pack.bytes)throw new Error("Invalid installer length");
   url=win.URL.createObjectURL(new win.Blob(blocks,{type:"application/octet-stream"}));urls.push(url);
   a=doc.createElement("a");a.href=url;a.download=pack.filename;doc.body.appendChild(a);a.click();a.remove();a=null;
   win.setTimeout(function(){release(url)},60000);notify(T.download_requested);return true;
  }catch(e){if(a)a.remove();if(url)release(url);notify(T.download_failure);return false}
  finally {busy=false;main.disabled=false}
 }
 function release(url){var i=urls.indexOf(url);if(i>=0){win.URL.revokeObjectURL(url);urls.splice(i,1)}}
 function render(){close();menu.replaceChildren();var methods=[["default",T.download_default],["baidu",T.download_baidu],["quark",T.download_quark]];
  methods.forEach(function(method){var available=method[0]==="default"?!!packages[product()]:!!safeURL(links()[method[0]]),b=doc.createElement("button");b.type="button";b.className="download-option";b.setAttribute("role","menuitem");b.dataset.channel=method[0];b.disabled=!available;b.textContent=method[1]+(available?"":"（"+(method[0]==="default"?T.download_unavailable_short:T.download_missing)+"）");b.addEventListener("click",function(){close();if(method[0]==="default")extract();else external(method[0])});menu.appendChild(b)});
  main.setAttribute("aria-label",T.download+" · "+(D.products.find(function(p){return p.id===product()})||{}).name);
 }
 function start(){close();if(packages[product()])extract();else if(safeURL(links().baidu))external("baidu");else if(safeURL(links().quark))external("quark");else notify(T.download_unconfigured)}
 function buildGuide(){
  if(built||!guide||!steps)return;built=true;
  (guide.steps||[]).forEach(function(step){var li=doc.createElement("li");li.textContent=step;steps.appendChild(li)});
  if(image&&guide.image_base64){image.src="data:image/png;base64,"+guide.image_base64;image.alt=guide.image_alt||"";image.hidden=false}
 }
 // The guide is advisory: without dialog support or guide data the download still starts.
 function confirmGuide(){if(!dialog||typeof dialog.showModal!=="function"||!guide){start();return}buildGuide();pending=start;dialog.showModal();if(proceed)proceed.focus({preventScroll:true})}
 function settle(accept){var action=pending;pending=null;if(dialog&&dialog.open)dialog.close();if(accept&&action)action()}
 main.addEventListener("click",confirmGuide);
 if(proceed)proceed.addEventListener("click",function(){settle(true)});
 var cancel=doc.getElementById("guide-dialog-cancel");
 if(cancel)cancel.addEventListener("click",function(){settle(false)});
 if(dialog)dialog.addEventListener("cancel",function(){pending=null});
 toggle.setAttribute("aria-label",T.download_methods);toggle.addEventListener("click",function(){menu.hidden=!menu.hidden;toggle.setAttribute("aria-expanded",String(!menu.hidden));if(!menu.hidden){positionMenu();var first=Array.from(menu.children).find(function(n){return !n.disabled});if(first)first.focus({preventScroll:true})}});
 doc.addEventListener("click",function(e){if(!e.target.closest(".download-split"))close()});
 doc.addEventListener("keydown",function(e){if(menu.hidden)return;if(e.key==="Escape"){close();toggle.focus();e.preventDefault()}else if(["ArrowDown","ArrowUp","Home","End"].indexOf(e.key)>=0&&e.target.closest(".download-split")){var items=Array.from(menu.children).filter(function(n){return !n.disabled});if(!items.length)return;var i=items.indexOf(doc.activeElement);i=e.key==="Home"?0:e.key==="End"?items.length-1:(i+(e.key==="ArrowUp"?-1:1)+items.length)%items.length;items[i].focus();e.preventDefault()}});
 win.addEventListener("resize",positionMenu);
 if(win.visualViewport){win.visualViewport.addEventListener("resize",positionMenu);win.visualViewport.addEventListener("scroll",positionMenu)}
 win.addEventListener("pagehide",function(){urls.slice().forEach(release);close()});render();
 return {render:render,extract:extract,close:close};
}
