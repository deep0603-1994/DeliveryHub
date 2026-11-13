(() => {
  'use strict';
  function write(client,table,values,{conflict,selection,single=false}={}) {
    let query=conflict?client.from(table).upsert(values,{onConflict:conflict,defaultToNull:false}):client.from(table).insert(values);
    if(selection)query=query.select(selection);
    return single?query.single():query;
  }
  function progress(processed,total) {
    const safeTotal=Math.max(1,Number(total)||1),safeProcessed=Math.min(safeTotal,Math.max(0,Number(processed)||0));
    return {processed:safeProcessed,total:safeTotal,percent:Math.round(safeProcessed/safeTotal*100)};
  }
  function yieldUi(background=false) {
    if(background&&document.visibilityState==='hidden')return new Promise(resolve=>{const channel=new MessageChannel();channel.port1.onmessage=()=>{channel.port1.close();channel.port2.close();resolve();};channel.port2.postMessage(null);});
    return new Promise(resolve=>requestAnimationFrame(()=>setTimeout(resolve,0)));
  }
  function errorMessage(error){return error?.message||String(error||'Unknown error');}
  window.deliveryImportCommon={write,progress,yieldUi,errorMessage};
})();