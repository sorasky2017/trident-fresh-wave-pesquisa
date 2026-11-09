/* Música de fundo gerada no próprio navegador (sem arquivos de áudio). Toca depois do primeiro toque na tela. */
(function(){
  var KEY='twfresh-musica-off',muted=false,ctx=null,master=null,timer=null,step=0,nextT=0,started=false;
  try{muted=localStorage.getItem(KEY)==='1'}catch(e){}
  var BPM=104,SPB=60/BPM,VOL=0.7;
  var chords=[[0,4,7],[9,12,16],[5,9,12],[7,11,14]];            // C, Am, F, G (semitons a partir de C)
  var roots=[0,-3,-7,-5];
  var penta=[0,2,4,7,9,12,14,16];
  var mel=[4,-1,2,-1,3,2,-1,1, 3,-1,5,-1,4,3,-1,2, 2,-1,4,-1,3,1,-1,0, 1,-1,3,-1,2,-1,-1,-1];
  function hz(semi,base){return (base||261.63)*Math.pow(2,semi/12)}
  function tone(t,f,d,type,g,a){
    var o=ctx.createOscillator(),v=ctx.createGain();o.type=type;o.frequency.value=f;
    v.gain.setValueAtTime(0.0001,t);v.gain.exponentialRampToValueAtTime(g,t+(a||0.02));v.gain.exponentialRampToValueAtTime(0.0001,t+d);
    o.connect(v);v.connect(master);o.start(t);o.stop(t+d+0.05);
  }
  function tick(t){
    var n=ctx.createBufferSource(),b=ctx.createBuffer(1,ctx.sampleRate*0.04,ctx.sampleRate),d=b.getChannelData(0);
    for(var i=0;i<d.length;i++)d[i]=(Math.random()*2-1)*(1-i/d.length);
    n.buffer=b;var v=ctx.createGain();v.gain.value=0.05;var f=ctx.createBiquadFilter();f.type='highpass';f.frequency.value=6000;
    n.connect(f);f.connect(v);v.connect(master);n.start(t);
  }
  function schedule(){
    if(!ctx)return;
    while(nextT<ctx.currentTime+0.35){
      var e=step%32,bar=Math.floor(step/8)%4,beat8=step%8,t=nextT;          // 8 colcheias por compasso
      if(beat8===0){chords[bar].forEach(function(s){tone(t,hz(s,130.81),SPB*4*0.98,'triangle',0.045,0.25)})}
      if(beat8%2===0){tone(t,hz(roots[bar],65.41),SPB*0.9,'sine',beat8===0?0.16:0.10,0.01)}
      if(beat8%2===1)tick(t);
      var m=mel[e];if(m>=0){tone(t,hz(penta[m%penta.length],523.25),SPB*0.9,'triangle',0.05,0.005)}
      step++;nextT+=SPB/2;
    }
    timer=setTimeout(schedule,120);
  }
  function start(){
    if(started||muted)return;
    try{
      ctx=ctx||new (window.AudioContext||window.webkitAudioContext)();
      if(!master){master=ctx.createGain();master.gain.value=VOL;master.connect(ctx.destination);window.__musica={ctx:ctx,master:master}}
      ctx.resume();started=true;nextT=ctx.currentTime+0.1;schedule();
    }catch(e){}
  }
  function stop(){started=false;clearTimeout(timer);if(ctx){try{ctx.suspend()}catch(e){}}}
  var btn=document.createElement('button');
  btn.setAttribute('aria-label','Ligar ou desligar a música');
  btn.style.cssText='position:fixed;right:12px;bottom:12px;z-index:99999;width:44px;height:44px;border-radius:50%;border:2px solid rgba(255,255,255,.9);background:rgba(9,44,80,.72);color:#fff;font-size:20px;line-height:1;cursor:pointer;box-shadow:0 4px 12px rgba(0,0,0,.25);display:grid;place-items:center;padding:0;-webkit-tap-highlight-color:transparent';
  function paint(){btn.textContent=muted?'🔇':'🔊'}
  btn.addEventListener('click',function(e){
    e.stopPropagation();muted=!muted;try{localStorage.setItem(KEY,muted?'1':'0')}catch(x){}
    if(muted)stop();else{started=false;start()}paint();
  });
  paint();
  function ready(){document.body.appendChild(btn)}
  if(document.body)ready();else document.addEventListener('DOMContentLoaded',ready);
  var first=function(){start();window.removeEventListener('pointerdown',first,true);window.removeEventListener('keydown',first,true)};
  window.addEventListener('pointerdown',first,true);window.addEventListener('keydown',first,true);
  document.addEventListener('visibilitychange',function(){if(muted||!ctx)return;if(document.hidden){try{ctx.suspend()}catch(e){}}else if(started){try{ctx.resume()}catch(e){}}});
})();
