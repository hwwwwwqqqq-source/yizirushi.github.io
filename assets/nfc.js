(() => {
 'use strict';
 let mappings=new Map(),request=0,controller=null,lastResult=null,socket=null,socketReady=false,panel=null,toast=null,cancelConnection=null;
 const VERSION='20261020-exhibit1',diagnostics=[];let diagnosticSerial=0;
 function selectedWord(){return window.YiziWordSelection?.getSelectedWord?.()??null}
 function focusState(){const el=document.activeElement;return {pageFocused:document.hasFocus?.()??null,element:el?.id||el?.tagName||'none'}}
 function trace(stage,data={}){diagnostics.push({sequence:++diagnosticSerial,time:new Date().toISOString(),stage,...data,focus:focusState()});if(diagnostics.length>160)diagnostics.shift();renderDiagnostics()}
 function diagnosticReport(){return {version:VERSION,reading,buffer,selected:selectedWord(),focus:focusState(),events:diagnostics}}
 function renderDiagnostics(){if(!panel)return;const out=panel.querySelector('[data-nfc-diagnostics]');if(out)out.textContent=JSON.stringify(diagnosticReport(),null,2)}
 const scriptURL=document.currentScript?.src||new URL('assets/nfc.js',location.href).href;
 function normalizeUID(uid){
  if(typeof uid!=='string'||uid.length>128)return '';
  const value=uid.trim().toUpperCase();
  if(/^[0-9A-F]+(?:[\s:-]+[0-9A-F]+)*$/.test(value))return value.replace(/[\s:-]/g,'');
  return /^[A-Z0-9_]{1,128}$/.test(value)?value:'';
 }
 function style(){if(document.querySelector('#nfc-styles'))return;const el=document.createElement('style');el.id='nfc-styles';el.textContent=`
 #nfc-reader{position:fixed;left:12px;bottom:6px;z-index:46;display:flex;gap:5px}#nfc-reader button{font:12px/16px SimSun,serif;padding:5px 9px;border:1px solid #98794c;border-radius:5px;background:#f0e2cbed;color:#573b22;cursor:pointer}#nfc-reader [aria-pressed=true]{background:#d7bf91}#nfc-test-panel{position:fixed;right:18px;bottom:18px;width:350px;max-width:calc(100vw - 36px);max-height:calc(100vh - 36px);overflow:auto;z-index:45;padding:19px;background:#f0e2cbee;color:#573b22;border:1px solid #98794c;border-radius:10px;font:14px/1.6 SimSun,serif;box-shadow:0 6px 22px #30271e33}
 #nfc-test-panel h2{font-size:18px;font-weight:400;margin:0 25px 10px 0}#nfc-test-panel p{margin:8px 0}#nfc-test-panel label{display:block;margin:8px 0 3px}#nfc-test-panel input{box-sizing:border-box;width:100%;font:14px/1.4 system-ui,sans-serif;padding:8px;background:#fff7e9;color:#4b351e;border:1px solid #a68d68;border-radius:4px}#nfc-test-panel button{font:inherit;color:#573b22;background:#ddc69d;border:1px solid #a68d68;border-radius:5px;padding:6px 12px;margin-top:10px}#nfc-test-panel .nfc-close{position:absolute;right:10px;top:5px;border:0;background:none;padding:2px 6px}#nfc-test-panel summary{cursor:pointer;margin-top:12px}#nfc-test-panel [role=status]{min-height:24px;margin-top:10px}#nfc-notice{position:fixed;right:20px;bottom:22px;z-index:44;padding:10px 17px;background:#f0e2cbef;border:1px solid #98794c;border-radius:6px;color:#573b22;font:15px/1.5 SimSun,serif;pointer-events:none}#nfc-notice[hidden]{display:none}
 /* Exhibition UI: keep the receiver and its handlers, hide all NFC controls/status. */
 #nfc-reader,#nfc-test-panel,#nfc-notice{display:none!important}`;document.head.append(el)}
 let toastTimer;
 function status(message){
  style();if(panel){if(toast)toast.hidden=true;panel.querySelector('[data-nfc-status]').textContent=message;return}
  if(!toast){toast=document.createElement('div');toast.id='nfc-notice';toast.setAttribute('role','status');document.body.append(toast)}
  toast.textContent=message;toast.hidden=false;clearTimeout(toastTimer);toastTimer=setTimeout(()=>{toast.hidden=true},4000);
 }
 function result(value){trace('result',value);lastResult=Object.freeze(value);window.dispatchEvent(new CustomEvent('yizi:nfc-result',{detail:lastResult,version:VERSION,focus:focusState()}));return lastResult}
 const ready=fetch(new URL('nfc-tags.json',scriptURL),{cache:'no-store'}).then(r=>{if(!r.ok)throw new Error('标签映射加载失败');return r.json()}).then(config=>{
  const next=new Map();if(!Array.isArray(config.mappings))throw new Error('标签映射格式错误');
  for(const row of config.mappings){const uid=normalizeUID(row.uid);if(!uid||next.has(uid)||!['xue','ru','xing'].includes(row.word)||typeof row.character!=='string')throw new Error('标签映射存在重复UID或无对应按钮');next.set(uid,Object.freeze({...row,uid}))}
  mappings=next;trace('mapping-ready',{count:next.size});return [...next.values()];
 });
 // Catch eagerly to avoid unhandled rejections; receiveUID still reports a load failure.
 ready.catch(()=>{});
 async function receiveUID(rawUID,{source='external'}={}){
  let serial=0;const uid=normalizeUID(rawUID);
  try{
   await ready;
   const mapping=mappings.get(uid);
   if(!mapping){status('未识别的标签');return result({ok:false,code:'unknown_uid',uid,message:'未识别的标签',source})}
   serial=++request;controller?.abort();controller=new AbortController();const signal=controller.signal;
   if(typeof window.YiziWordSelection?.click!=='function')throw new Error('原字卡切换函数尚未初始化');
   trace('switch-start',{uid,word:mapping.word,source});
   await window.YiziWordSelection.click(mapping.word,{signal});
   if(selectedWord()!==mapping.word)throw new Error('切换函数未更新对应字形状态');
   if(serial!==request||signal.aborted)return {ok:false,code:'superseded',uid};
   status('已选择「'+mapping.character+'」');return result({ok:true,uid,word:mapping.word,character:mapping.character,source});
  }catch(error){
   if(error.name==='AbortError'||(serial&&serial!==request))return {ok:false,code:'superseded',uid};
   const message=error.message||'标签选择失败';status(message);return result({ok:false,uid,code:'selection_failed',message,source});
  }
 }
 // One frame accumulator for keyboard, native input, paste and IME commit.
 // The device's registered frames are exactly ten decimal digits, no Enter.
 let reading=false,buffer='',scanTimer=null,scanQueue=Promise.resolve(),dock=null,scanCount=0,receiver=null;
 const SCAN_IDLE_MS=200,PARTIAL_IDLE_MS=1200,UID_LENGTH=10;
 function clearScan(){clearTimeout(scanTimer);scanTimer=null;buffer='';if(receiver)receiver.value=''}
 function focusReceiver(){if(reading&&receiver&&!document.hidden)try{receiver.focus({preventScroll:true})}catch(error){trace('focus-failed',{message:error.message})}}
 function enqueue(uid,source){
  const scanId=++scanCount;trace('scan-ended',{scanId,raw:uid,uid,source,bufferAfter:buffer});
  const task=scanQueue.then(async()=>{
   trace('processing',{scanId,uid,source});
   const response=await receiveUID(uid,{source});
   trace('switch-finished',{scanId,uid,response,selected:selectedWord()});
   focusReceiver();return response;
  });
  scanQueue=task.catch(error=>{trace('queue-error',{scanId,message:error.message});status('读卡处理失败：'+error.message)});return task;
 }
 function finishPartial(){const uid=buffer;clearScan();if(reading&&uid)void enqueue(uid,'incomplete-frame')}
 function schedulePartial(){clearTimeout(scanTimer);scanTimer=setTimeout(()=>{
  trace('idle-incomplete',{raw:buffer,length:buffer.length});
  scanTimer=setTimeout(finishPartial,PARTIAL_IDLE_MS-SCAN_IDLE_MS);
 },SCAN_IDLE_MS)}
 function appendDigits(value,source){
  if(String(value).length>128){clearScan();trace('input-too-long',{source,length:String(value).length});status('未识别的标签');return}
  const text=String(value).replace(/[０-９]/g,c=>String(c.charCodeAt(0)-0xff10)).trim();
  trace('raw-input',{source,raw:value});
  if(!/^\d+$/.test(text)){clearScan();if(text)void enqueue(text,source);return}
  for(const d of text){
   if(!buffer)trace('scan-start',{source,bufferBefore:''});
   buffer+=d;
   if(buffer.length===UID_LENGTH){const uid=buffer;clearScan();void enqueue(uid,source)}
  }
  if(buffer)schedulePartial();renderDiagnostics();
 }
 function setReading(enabled){
  reading=Boolean(enabled);clearScan();
  if(dock){const button=dock.querySelector('[data-reader-toggle]');button.textContent=reading?'停用读卡':'启用读卡';button.setAttribute('aria-pressed',String(reading));dock.dataset.enabled=String(reading)}
  trace('reader-state',{reading});status(reading?'读卡已启用，请在此网页刷卡':'读卡已停用');focusReceiver();return reading;
 }
 function keyboardDigit(event){
  if(/^[0-9]$/.test(event.key))return event.key;
  const code=/^(?:Digit|Numpad)([0-9])$/.exec(event.code||'');if(code)return code[1];
  const n=event.keyCode||event.which;if(n>=48&&n<=57)return String(n-48);if(n>=96&&n<=105)return String(n-96);return null;
 }
 window.addEventListener('keydown',event=>{
  if(!reading)return;
  const target=event.target,readerField=target===receiver||target?.id==='nfc-uid';
  if(!readerField&&target?.closest?.('input,textarea,select,[contenteditable]:not([contenteditable="false"])'))return;
  if(event.ctrlKey||event.altKey||event.metaKey)return;
  trace('key',{key:event.key,code:event.code,keyCode:event.keyCode,composing:event.isComposing,cancelable:event.cancelable});
  // IME/non-cancelable input is read from the actual textbox value at commit.
  // Do not discard it or count both keydown and the subsequent native input.
  if(event.isComposing||event.keyCode===229||event.cancelable===false)return;
  const digit=keyboardDigit(event);
  if(digit!==null){event.preventDefault();appendDigits(digit,'keyboard');return}
  // NumLock/Shift/Tab/etc. must not erase a complete or partial reader frame.
  if(event.key==='Escape'&&buffer){trace('scan-cancelled',{raw:buffer});clearScan()}
 },true);
 function receiveNativeInput(event){
  if(!reading||event.isComposing)return;
  const input=event.target,value=input.value;if(!value)return;
  input.value='';appendDigits(value,event.type==='compositionend'?'composition':'input');
 }
 window.addEventListener('blur',()=>{trace('window-blur',{discardedPartial:buffer});clearScan()});
 window.addEventListener('focus',()=>{trace('window-focus');focusReceiver()});
 document.addEventListener('visibilitychange',()=>{trace('visibility',{hidden:document.hidden});if(document.hidden)clearScan()});
 document.addEventListener('focusin',()=>renderDiagnostics());
 function mountReader(){
  style();receiver=document.createElement('input');receiver.id='nfc-receiver';receiver.type='text';receiver.tabIndex=-1;
  receiver.setAttribute('aria-label','NFC读卡接收');receiver.setAttribute('inputmode','numeric');receiver.setAttribute('autocomplete','off');receiver.setAttribute('maxlength','128');receiver.setAttribute('lang','en');receiver.spellcheck=false;
  receiver.style.cssText='position:fixed;left:-10000px;top:0;width:1px;height:1px;opacity:0;pointer-events:none';
  receiver.addEventListener('input',receiveNativeInput);receiver.addEventListener('compositionend',receiveNativeInput);document.body.append(receiver);
  dock=document.createElement('div');dock.id='nfc-reader';dock.dataset.enabled='false';dock.dataset.version=VERSION;
  dock.innerHTML='<button type="button" data-reader-toggle aria-pressed="false">启用读卡</button><button type="button" data-reader-test>编号测试</button>';
  document.body.append(dock);dock.querySelector('[data-reader-toggle]').onclick=()=>setReading(!reading);dock.querySelector('[data-reader-test]').onclick=showTestMode;
  trace('initialized',{version:VERSION});
 }
 function disconnect(){const old=socket,cancel=cancelConnection;cancelConnection=null;socket=null;socketReady=false;cancel?.();if(old)old.close();}
 function connect({url='ws://127.0.0.1:8765',token=''}={}){
  let endpoint;try{endpoint=new URL(url)}catch{return Promise.reject(new Error('桥接地址无效'))}
  if(!['ws:','wss:'].includes(endpoint.protocol)||!['127.0.0.1','localhost','[::1]'].includes(endpoint.hostname)||endpoint.username||endpoint.password||endpoint.search)return Promise.reject(new Error('请使用本机ws/wss地址，密钥不要放入URL'));
  disconnect();return new Promise((resolve,reject)=>{
   let ws,settled=false;try{ws=new WebSocket(endpoint.href)}catch(error){return reject(error)}socket=ws;
   const timer=setTimeout(()=>{if(!settled){settled=true;if(socket===ws){cancelConnection=null;disconnect()}reject(new Error('桥接连接超时'))}},6000);
   cancelConnection=()=>{clearTimeout(timer);if(!settled){settled=true;reject(new Error('桥接连接已取消'))}};
   function fail(){if(socket!==ws)return;socketReady=false;if(!settled){settled=true;clearTimeout(timer);reject(new Error('桥接未连接，请检查程序、密钥或浏览器本地网络限制'))}else status('桥接已断开')}
   ws.onopen=()=>{if(socket===ws)ws.send(JSON.stringify({type:'auth',token}))};
   ws.onmessage=async event=>{
    if(socket!==ws||typeof event.data!=='string'||event.data.length>4096)return;
    let data;try{data=JSON.parse(event.data)}catch{return}
    if(data.type==='ready'){socketReady=true;settled=true;cancelConnection=null;clearTimeout(timer);status('外部桥接已连接');resolve();return}
    if(data.type==='uid'&&socketReady){const response=await receiveUID(data.uid,{source:'websocket'});if(socket===ws&&ws.readyState===WebSocket.OPEN)ws.send(JSON.stringify({type:'result',id:data.id,...response}))}
   };ws.onerror=fail;ws.onclose=fail;
  });
 }
 function showTestMode(){
  if(panel){panel.querySelector('#nfc-uid').focus();return}style();panel=document.createElement('aside');panel.id='nfc-test-panel';panel.setAttribute('aria-label','RFID/NFC标签测试');
  panel.innerHTML='<h2>RFID / NFC 标签测试</h2><p>接收版本：20261020-exhibit1</p><button class="nfc-close" aria-label="关闭标签测试">×</button><p data-nfc-mappings>正在读取映射…</p><form><label for="nfc-uid">标签 UID</label><input id="nfc-uid" name="uid" maxlength="128" autocomplete="off" inputmode="numeric" placeholder="例如 1954379268"><button type="submit">测试标签</button></form><div role="status" aria-live="polite" data-nfc-status>输入编号并点击测试；启用读卡后可自动接收</div><details data-nfc-debug><summary>读卡诊断（仅本页记录）</summary><button data-nfc-export type="button">下载诊断</button><pre data-nfc-diagnostics style="white-space:pre-wrap;overflow-wrap:anywhere;max-height:240px;overflow:auto;font:11px/1.5 monospace"></pre></details><details><summary>外部读卡程序连接</summary><label for="nfc-bridge-url">本机桥接地址</label><input id="nfc-bridge-url" value="ws://127.0.0.1:8765" autocomplete="off"><label for="nfc-bridge-token">桥接密钥</label><input id="nfc-bridge-token" type="password" autocomplete="off"><button data-nfc-connect type="button">连接桥接</button><button data-nfc-disconnect type="button">断开</button></details>';
  document.body.append(panel);panel.querySelector('.nfc-close').onclick=()=>{panel.remove();panel=null;dock.querySelector('[data-reader-toggle]').focus()};
  const input=panel.querySelector('#nfc-uid');
  // Paste/autofill generates input rather than keydown. Keydown already
  // prevents native insertion, so this cannot double-count keyboard digits.
  input.oninput=receiveNativeInput;input.oncompositionend=receiveNativeInput;
  renderDiagnostics();
  panel.querySelector('[data-nfc-export]').onclick=()=>{
   try{const url=URL.createObjectURL(new Blob([JSON.stringify(diagnosticReport(),null,2)],{type:'application/json'}));const link=document.createElement('a');link.href=url;link.download='yizi-nfc-diagnostics.json';link.click();setTimeout(()=>URL.revokeObjectURL(url),1000)}catch(error){status('诊断下载失败，可复制下方记录：'+error.message)}
  };
  panel.querySelector('form').onsubmit=event=>{event.preventDefault();const uid=input.value;clearScan();input.value='';if(uid)void enqueue(uid,'test')};
  panel.querySelector('[data-nfc-connect]').onclick=()=>connect({url:panel.querySelector('#nfc-bridge-url').value,token:panel.querySelector('#nfc-bridge-token').value}).catch(e=>status(e.message));
  panel.querySelector('[data-nfc-disconnect]').onclick=()=>{disconnect();status('桥接已断开')};
  ready.then(rows=>{if(panel)panel.querySelector('[data-nfc-mappings]').textContent=rows.map(r=>r.uid+' → '+r.character).join('；')}).catch(e=>status(e.message));
 }
 window.addEventListener('message',event=>{
  // Only this page's own window / trusted injected bridge may forward messages.
  if(event.source!==window||event.origin!==location.origin||event.data?.type!=='YIZI_NFC_UID')return;
  void receiveUID(event.data.uid,{source:'postMessage'});
 });
 window.YiziNFC=Object.freeze({version:VERSION,receiveUID,showTestMode,setReading,getDiagnostics:()=>diagnostics.map(row=>({...row})),connect,disconnect,normalizeUID,ready,getMappings:()=>[...mappings.values()].map(r=>({...r})),getState:()=>({selectedWord:selectedWord(),connected:socketReady,reading,bufferLength:buffer.length,scanCount,lastResult,version:VERSION,focus:focusState()})});
 mountReader();
 // A keyboard-emulating reader needs no browser hardware permission.
 // Enable the existing receiver automatically; focus is also restored on window focus.
 setReading(true);
 if(new URLSearchParams(location.search).get('nfc-test')==='1')showTestMode();
})();
