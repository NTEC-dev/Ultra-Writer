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
  const contentArea = document.getElementById('content');
  const getValue = (id) => contentArea.querySelector(`#${id}`)?.value || '';

  const title = getValue('title');
  const filename = `${title.replace(/[^a-z0-9]/gi, '_') || 'Project-Triage'}.doc`;

  const formatText = (text) => text.replace(/\n/g, '<br>');

  const content = `
    <h1>Project Triage: ${title}</h1>
    <hr>
    
    <h2>Decision</h2>
    <table border="1" cellpadding="5" style="border-collapse: collapse; width: 100%;">
        <tr><td width="30%"><strong>Project FP No.</strong></td><td>${getValue('fpno')}</td></tr>
        <tr><td><strong>Owner</strong></td><td>${getValue('owner')}</td></tr>
        <tr><td><strong>Stage</strong></td><td>${getValue('stage')}</td></tr>
        <tr><td><strong>Sponsor / Client</strong></td><td>${getValue('sponsor')}</td></tr>
        <tr><td><strong>Start Date</strong></td><td>${getValue('startDate')}</td></tr>
        <tr><td><strong>End Date</strong></td><td>${getValue('endDate')}</td></tr>
    </table>

    <h3>Executive Summary</h3>
    <p>${formatText(getValue('exec'))}</p>

    <h3>Why participate (need/opportunity & strategic fit)</h3>
    <p>${formatText(getValue('why'))}</p>

    <h3>What we will deliver (scope, outputs, success)</h3>
    <p>${formatText(getValue('deliver'))}</p>
    
    <h3>Top risks & dependencies</h3>
    <table border="1" cellpadding="5" style="border-collapse: collapse; width: 100%;">
        <tr><th>Risk</th><th>Mitigation</th></tr>
        <tr><td>${formatText(getValue('risk1'))}</td><td>${formatText(getValue('mit1'))}</td></tr>
        <tr><td>${formatText(getValue('risk2'))}</td><td>${formatText(getValue('mit2'))}</td></tr>
        <tr><td>${formatText(getValue('risk3'))}</td><td>${formatText(getValue('mit3'))}</td></tr>
    </table>
    
    <h3>Next steps & owners</h3>
    <table border="1" cellpadding="5" style="border-collapse: collapse; width: 100%;">
        <tr><th>Action</th><th>Owner</th><th>Due</th></tr>
        <tr><td>${formatText(getValue('act1'))}</td><td>${getValue('own1')}</td><td>${getValue('due1')}</td></tr>
        <tr><td>${formatText(getValue('act2'))}</td><td>${getValue('own2')}</td><td>${getValue('due2')}</td></tr>
        <tr><td>${formatText(getValue('act3'))}</td><td>${getValue('own3')}</td><td>${getValue('due3')}</td></tr>
    </table>
    
    <br/>
    <hr>

    <h2>Plan</h2>
    <h3>Overall approach</h3>
    <p>${formatText(getValue('plan_approach'))}</p>
    <h3>Key milestones</h3>
    <p>${formatText(getValue('plan_milestones'))}</p>

    <br/>
    <hr>
    
    <h2>Risks</h2>
    <h3>Risk register</h3>
    <p>${formatText(getValue('risks_register'))}</p>
    <h3>Top issues</h3>
    <table border="1" cellpadding="5" style="border-collapse: collapse; width: 100%;">
        <tr><th>Issue</th><th>Owner</th></tr>
        <tr><td>${formatText(getValue('issues_i1'))}</td><td>${getValue('issues_o1')}</td></tr>
        <tr><td>${formatText(getValue('issues_i2'))}</td><td>${getValue('issues_o2')}</td></tr>
        <tr><td>${formatText(getValue('issues_i3'))}</td><td>${getValue('issues_o3')}</td></tr>
    </table>
  `;

  const preHtml = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML To Doc</title></head><body>";
  const postHtml = "</body></html>";
  const html = preHtml + content + postHtml;

  const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
  const url = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(html);
  const downloadLink = document.createElement("a");
  document.body.appendChild(downloadLink);
  
  if (navigator.msSaveOrOpenBlob) {
      navigator.msSaveOrOpenBlob(blob, filename);
  } else {
      downloadLink.href = url;
      downloadLink.download = filename;
      downloadLink.click();
  }
  
  document.body.removeChild(downloadLink);
}

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
