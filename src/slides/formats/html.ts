import { escapeHtml } from '@/lib/utils';
import type { Presentation } from '../model';
import { buildOf } from '../show/engine';
import slidesCss from '../slides.css?inline';

/** A self-contained web page that plays the presentation (keyboard, click, swipe, fullscreen). */
export async function exportHtmlShow(pres: Presentation, title: string): Promise<Uint8Array> {
  const { slidesHtml } = await import('./raster');
  const visible = pres.slides.map((s, i) => (s.hidden ? -1 : i)).filter((i) => i >= 0);
  const { slides, fontCss } = await slidesHtml(pres, visible);
  const data = visible.map((i) => {
    const s = pres.slides[i];
    const b = buildOf(s);
    const item = (it: (typeof b.steps)[number][number]) => ({ el: it.anim.el, cls: it.anim.cls, effect: it.anim.effect, dir: it.anim.dir ?? null, dur: it.anim.dur, at: it.offset, para: it.para ?? null, frame: (() => {
      const e = s.elements.find((x) => x.id === it.anim.el);
      return e ? [e.x, e.y, e.w, e.h] : [0, 0, 0, 0];
    })() });
    return { t: s.transition && s.transition.type !== 'none' ? { type: s.transition.type, dir: s.transition.dir ?? null, dur: s.transition.dur ?? 700, after: s.transition.after ?? null } : null, auto: b.auto?.map(item) ?? [], steps: b.steps.map((st) => st.map(item)), notes: s.notes ?? '' };
  });
  const name = escapeHtml(title.replace(/\.[^.]+$/, ''));
  const html = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${name}</title>
<meta name="generator" content="Affice">
<style>
${fontCss}
${slidesCss}
html,body{margin:0;height:100%;background:#000;overflow:hidden;font-family:system-ui,sans-serif}
#deck{position:fixed;inset:0;display:flex;align-items:center;justify-content:center}
#stage{position:relative;overflow:hidden;width:${pres.size.w}px;height:${pres.size.h}px;transform-origin:center center;flex:none}
.slide{position:absolute;inset:0;display:none}
.slide.on,.slide.leaving{display:block}
.slide .sl-slide{pointer-events:auto}
#bar{position:fixed;left:0;bottom:0;height:3px;background:#f26a26;transition:width .3s ease;z-index:5}
#hud{position:fixed;right:16px;bottom:14px;color:#ccc;font:13px system-ui,sans-serif;background:rgba(20,20,26,.6);padding:6px 10px;border-radius:99px;opacity:0;transition:opacity .3s;z-index:5;user-select:none}
body.hud #hud{opacity:1}
#hud button{background:none;border:0;color:#fff;font:inherit;cursor:pointer;padding:0 6px}
.x-hidden{visibility:hidden!important}
</style></head>
<body>
<div id="deck"><div id="stage">
${slides.map((s, i) => `<section class="slide" data-i="${i}">${s}</section>`).join('\n')}
</div></div>
<div id="bar"></div>
<div id="hud"><button id="prev" aria-label="Previous">‹</button><span id="count"></span><button id="next" aria-label="Next">›</button><button id="fs" aria-label="Full screen">⛶</button></div>
<script>
(function(){
var D=${JSON.stringify(data)};
var W=${pres.size.w},H=${pres.size.h};
var stage=document.getElementById('stage'),els=[].slice.call(document.querySelectorAll('.slide'));
var cur=0,done=0,busy=false,timer=0,hudT=0;
function fit(){var s=Math.min(innerWidth/W,innerHeight/H);stage.style.transform='scale('+s+')';}
addEventListener('resize',fit);fit();
function q(sec,id){return sec.querySelector('[data-el="'+id+'"]');}
function target(sec,it){var n=q(sec,it.el);if(!n)return null;if(it.para!==null){var ps=n.querySelectorAll('.sl-p');return ps[it.para]||null;}return n;}
function initial(i){var sec=els[i],d=D[i],seen={};
  [].concat(d.auto,[].concat.apply([],d.steps)).forEach(function(it){var k=it.el+'#'+it.para;if(seen[k])return;seen[k]=1;var t=target(sec,it);if(t&&it.cls==='entr')t.classList.add('x-hidden');});}
function frames(it,base){var b=base?base+' ':'',f=it.frame,dir=it.dir||'d',off;
  if(dir==='u')off='translate(0px,'+(-(f[1]+f[3]+20))+'px)';else if(dir==='l')off='translate('+(-(f[0]+f[2]+20))+'px,0px)';else if(dir==='r')off='translate('+(W-f[0]+20)+'px,0px)';else off='translate(0px,'+(H-f[1]+20)+'px)';
  switch(it.effect){
    case 'fly':return [{transform:off+' '+b},{transform:base||'none'}];
    case 'float':return [{opacity:0,transform:'translate(0px,'+(dir==='d'?-60:60)+'px) '+b},{opacity:1,transform:base||'none'}];
    case 'zoom':return [{opacity:0,transform:b+'scale(.3)'},{opacity:1,transform:b+'scale(1)'}];
    case 'wipe':var c=dir==='u'?'inset(0 0 100% 0)':dir==='l'?'inset(0 100% 0 0)':dir==='r'?'inset(0 0 0 100%)':'inset(100% 0 0 0)';return [{clipPath:c},{clipPath:'inset(0 0 0 0)'}];
    case 'split':return [{clipPath:'inset(0 50% 0 50%)'},{clipPath:'inset(0 0% 0 0%)'}];
    case 'wheel':return [{opacity:0,transform:b+'rotate(-200deg) scale(.2)'},{opacity:1,transform:b+'rotate(0deg) scale(1)'}];
    case 'bounce':return [{opacity:0,transform:b+'translate(0,-160px)',offset:0},{opacity:1,transform:b+'translate(0,0)',offset:.55},{transform:b+'translate(0,-40px)',offset:.72},{transform:b+'translate(0,0)',offset:.86},{transform:b+'translate(0,-10px)',offset:.93},{transform:b+'translate(0,0)',offset:1}];
    case 'pulse':return [{transform:b+'scale(1)'},{transform:b+'scale(1.1)'},{transform:b+'scale(1)'}];
    case 'spin':return [{transform:b+'rotate(0deg)'},{transform:b+'rotate(360deg)'}];
    case 'grow':return [{transform:b+'scale(1)'},{transform:b+'scale(1.35)'},{transform:b+'scale(1)'}];
    case 'teeter':return [0,-6,6,-6,6,0].map(function(d){return {transform:b+'rotate('+d+'deg)'};});
    case 'transparency':return [{opacity:1},{opacity:.35},{opacity:1}];
    default:return [{opacity:0},{opacity:1}];
  }}
function play(items){var sec=els[cur];items.forEach(function(it){var t=target(sec,it);if(!t)return;var base=it.para!==null?'':t.style.transform;var f=frames(it,base);
  if(it.cls==='exit')f=f.slice().reverse();t.classList.remove('x-hidden');
  t.animate(f,{duration:Math.max(1,it.effect==='appear'||it.effect==='disappear'?1:it.dur),delay:it.at,easing:it.cls==='exit'?'cubic-bezier(.6,0,.8,.3)':'cubic-bezier(.2,.7,.2,1)',fill:it.cls==='emph'?'none':'both'});});}
function trans(t){var dir=t.dir,k;function v(d,s){s=s?-1:1;return d==='r'?[(-100*s)+'%','0%']:d==='u'?['0%',(100*s)+'%']:d==='d'?['0%',(-100*s)+'%']:[(100*s)+'%','0%'];}
  switch(t.type){case 'push':var a=v(dir),o=v(dir,1);return {i:[{transform:'translate('+a[0]+','+a[1]+')'},{transform:'none'}],o:[{transform:'none'},{transform:'translate('+o[0]+','+o[1]+')'}]};
  case 'cover':a=v(dir);return {i:[{transform:'translate('+a[0]+','+a[1]+')'},{transform:'none'}]};
  case 'reveal':o=v(dir,1);return {o:[{transform:'none'},{transform:'translate('+o[0]+','+o[1]+')',opacity:.4}],top:1};
  case 'wipe':return {i:[{clipPath:dir==='r'?'inset(0 100% 0 0)':dir==='u'?'inset(100% 0 0 0)':dir==='d'?'inset(0 0 100% 0)':'inset(0 0 0 100%)'},{clipPath:'inset(0 0 0 0)'}]};
  case 'split':return {i:[{clipPath:dir==='u'?'inset(50% 0 50% 0)':'inset(0 50% 0 50%)'},{clipPath:'inset(0 0 0 0)'}]};
  case 'circle':return {i:[{clipPath:'circle(0% at 50% 50%)'},{clipPath:'circle(75% at 50% 50%)'}]};
  case 'zoom':return {i:[{opacity:0,transform:'scale(.4)'},{opacity:1,transform:'none'}],o:[{opacity:1},{opacity:0,transform:'scale(1.3)'}]};
  case 'dissolve':return {i:[{opacity:0,filter:'blur(8px)'},{opacity:1,filter:'blur(0)'}]};
  default:return {i:[{opacity:0},{opacity:1}]};}}
function show(i,animate){if(i<0||i>=els.length||busy)return;var prev=cur;cur=i;done=0;initial(i);
  els.forEach(function(s,k){s.classList.toggle('on',k===i);s.classList.remove('leaving');s.style.zIndex='';});
  var t=D[i].t;if(animate&&t&&prev!==i){var k=trans(t),inn=els[i],out=els[prev];out.classList.add('leaving');out.style.zIndex=k.top?2:0;inn.style.zIndex=1;busy=true;
    if(k.i)inn.animate(k.i,{duration:t.dur,easing:'cubic-bezier(.4,0,.2,1)'});if(k.o)out.animate(k.o,{duration:t.dur,easing:'cubic-bezier(.4,0,.2,1)',fill:'forwards'});
    setTimeout(function(){busy=false;out.classList.remove('leaving');out.getAnimations().forEach(function(a){a.cancel();});if(D[i].auto.length)play(D[i].auto);},t.dur);}
  else if(D[i].auto.length)play(D[i].auto);
  document.getElementById('count').textContent=(i+1)+' / '+els.length;document.getElementById('bar').style.width=(100*(i+1)/els.length)+'%';
  clearTimeout(timer);if(t&&t.after)timer=setTimeout(next,t.after);location.hash='#'+(i+1);}
function next(){if(busy)return;var d=D[cur];if(done<d.steps.length){play(d.steps[done++]);return;}if(cur<els.length-1)show(cur+1,true);}
function prev(){if(busy)return;if(done>0){done--;initial(cur);for(var s=0;s<done;s++)D[cur].steps[s].forEach(function(it){var t=target(els[cur],it);if(t)t.classList.toggle('x-hidden',it.cls==='exit');});return;}if(cur>0){show(cur-1,false);var d=D[cur];done=d.steps.length;d.steps.forEach(function(st){st.forEach(function(it){var t=target(els[cur],it);if(t)t.classList.toggle('x-hidden',it.cls==='exit');});});}}
addEventListener('keydown',function(e){var k=e.key;if(['ArrowRight','ArrowDown','PageDown',' ','Enter','n'].indexOf(k)>=0){e.preventDefault();next();}else if(['ArrowLeft','ArrowUp','PageUp','Backspace','p'].indexOf(k)>=0){e.preventDefault();prev();}else if(k==='Home')show(0);else if(k==='End')show(els.length-1);else if(k==='f'||k==='F')fs();});
document.getElementById('deck').addEventListener('click',function(e){var a=e.target.closest&&e.target.closest('a[href]');if(a){var h=a.getAttribute('href');if(/^#slide:\\d+$/.test(h)){show(parseInt(h.slice(7),10)-1,true);e.preventDefault();return;}if(/^(https?:|mailto:)/.test(h)){window.open(h,'_blank','noopener');e.preventDefault();return;}}next();});
var tx=null;addEventListener('touchstart',function(e){tx=e.touches[0].clientX;},{passive:true});addEventListener('touchend',function(e){if(tx===null)return;var dx=e.changedTouches[0].clientX-tx;if(Math.abs(dx)>40)(dx<0?next:prev)();tx=null;});
function fs(){if(document.fullscreenElement)document.exitFullscreen();else document.documentElement.requestFullscreen&&document.documentElement.requestFullscreen();}
document.getElementById('next').onclick=function(e){e.stopPropagation();next();};document.getElementById('prev').onclick=function(e){e.stopPropagation();prev();};document.getElementById('fs').onclick=function(e){e.stopPropagation();fs();};
addEventListener('mousemove',function(){document.body.classList.add('hud');clearTimeout(hudT);hudT=setTimeout(function(){document.body.classList.remove('hud');},2200);});
var start=parseInt((location.hash||'#1').slice(1),10)-1;show(isNaN(start)?0:Math.max(0,Math.min(els.length-1,start)),false);
})();
</script>
</body></html>`;
  return new TextEncoder().encode(html);
}
