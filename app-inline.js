(function(){
function switchTab(id){
  var tabs = document.querySelectorAll('.tabs button');
  for(var i=0;i<tabs.length;i++){ tabs[i].classList.remove('active'); }
  var btn = document.querySelector('.tabs button[data-tab="'+id+'"]');
  if(btn) btn.classList.add('active');
  var panels = document.querySelectorAll('.tab-panel');
  for(var j=0;j<panels.length;j++){ panels[j].classList.remove('active'); }
  var pane = document.getElementById('tab-'+id);
  if(pane) pane.classList.add('active');
}
window.switchTab = switchTab;

function _scope(){ return document.getElementById('content') || document; }
function _get(id){ var el=_scope().querySelector('#'+id); if(!el) return ''; if(el.type==='checkbox') return !!el.checked; return el.value||''; }
function _set(id,v){ var el=_scope().querySelector('#'+id); if(!el) return; if(el.type==='checkbox') el.checked=!!v; else el.value=v; }
function _downloadBlob(blob,name){ var a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name; document.body.appendChild(a); a.click(); URL.revokeObjectURL(a.href); a.remove(); }

window.printPDF = function(){ var prev=document.title; var title=_get('title')||'Project'; var fp=_get('fpno')||''; document.title='Project-Triage-'+(fp||title); window.print(); setTimeout(function(){ document.title=prev; }, 500); };

window.exportRTF = function(){
  function escRtf(s){return String(s||'').replace(/[\\{}]/g,function(m){return m==='\\'?'\\\\':'\\'+m;}).replace(/\n/g,'\\line ');} 
  function sanitize(s){return String(s||'').replace(/[\u2013\u2014\u2192]/g,'-');} 
  function rtfP(t){return "{\\pard "+escRtf(sanitize(t))+"\\par}\n";} 
  function rtfH2(t){return "{\\pard\\sa240\\b "+escRtf(t)+"\\par}\n";} 
  function rtfLabelVal(l,v){return "{\\pard\\b "+escRtf(l+": ")+"\\b0 "+escRtf(sanitize(v||"—"))+"\\par}\n";} 
  var r="{\\rtf1\\ansi "; 
  r+="{\\pard\\qc\\sa240\\b Project Triage\\par}\n"; 
  r+=rtfH2("Decision"); 
  r+=rtfLabelVal("Project title",_get('title')); 
  r+=rtfLabelVal("Project FP No.",_get('fpno')); 
  r+=rtfLabelVal("Owner",_get('owner')); 
  var stageEl=_scope().querySelector('#stage'); r+=rtfLabelVal("Stage",stageEl?stageEl.value:''); 
  r+=rtfLabelVal("Sponsor / Client",_get('sponsor')); 
  r+=rtfLabelVal("Start date",_get('start')); 
  r+=rtfLabelVal("End date",_get('end')); 
  r+=rtfH2("Executive summary"); r+=rtfP(_get('exec')); 
  r+=rtfH2("Why participate"); r+=rtfP(_get('why')); 
  _downloadBlob(new Blob([r],{type:'application/rtf'}),'Project-Triage.rtf'); 
};

window.exportDOCX = function(){
  if(window.PizZip && window.docxtemplater){
    fetch('WRC-General.dotx').then(function(r){return r.arrayBuffer();}).then(function(buf){
      try{
        var zip=new PizZip(buf);
        var doc=new window.docxtemplater(zip,{paragraphLoop:true,linebreaks:true});
        var data={}; var els=_scope().querySelectorAll('input,textarea,select');
        for(var i=0;i<els.length;i++){ var el=els[i]; if(el.id){ data[el.id]=(el.type==='checkbox')?el.checked:el.value; } }
        doc.setData(data); doc.render();
        _downloadBlob(doc.getZip().generate({type:'blob'}),'Project-Triage.docx');
      }catch(e){ alert('DOCX template merge failed: '+e.message); }
    }).catch(function(){ alert('Could not read WRC-General.dotx.'); });
  }else if(window.htmlDocx && window.htmlDocx.asBlob){
    var html='<!doctype html><html><head><meta charset="utf-8"></head><body>'+_scope().innerHTML+'</body></html>';
    _downloadBlob(window.htmlDocx.asBlob(html),'Project-Triage.docx');
  }else{
    alert('DOCX export needs either PizZip+Docxtemplater (template) or html-docx.js (HTML→DOCX).');
  }
};

window.saveForm = function(){
  var data={}; var els=_scope().querySelectorAll('input,textarea,select');
  for(var i=0;i<els.length;i++){ var el=els[i]; if(el.id){ data[el.id]=(el.type==='checkbox')?el.checked:el.value; } }
  try{ localStorage.setItem('projectForms.v3h.triage', JSON.stringify(data)); alert('Saved.'); }catch(e){ alert('Save failed: '+e.message); }
};
window.loadForm = function(){
  var raw=localStorage.getItem('projectForms.v3h.triage');
  if(!raw){ alert('No saved data.'); return; }
  var data; try{ data=JSON.parse(raw); }catch(e){ alert('Corrupt saved data'); return; }
  for(var k in data){ if(Object.prototype.hasOwnProperty.call(data,k)) _set(k,data[k]); }
  alert('Loaded.');
};

// Default tab
document.addEventListener('DOMContentLoaded', function(){ switchTab('decision'); });
})();