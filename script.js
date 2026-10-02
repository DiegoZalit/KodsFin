const state={
  workbook:null,fileName:"KodsFin.xlsx",fileHandle:null,
  lancamentos:[],contas:[],receitas:[],categorias:[],
  currentView:"dashboard", charts:{}
};

const $=s=>document.querySelector(s);
const money=v=>new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"}).format(Number(v)||0);
const todayISO=()=>new Date().toISOString().slice(0,10);
const monthKey=d=>String(d||"").slice(0,7);
const uid=()=>Date.now().toString(36)+"-"+Math.random().toString(36).slice(2,8);
const esc=s=>String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m]));

function toast(msg){const t=$("#toast");t.textContent=msg;t.classList.add("show");setTimeout(()=>t.classList.remove("show"),2200)}
function normalizeDate(v){
  if(!v)return "";
  if(typeof v==="number" && XLSX?.SSF){const d=XLSX.SSF.parse_date_code(v);return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`}
  const s=String(v); if(/^\\d{4}-\\d{2}-\\d{2}/.test(s))return s.slice(0,10);
  const m=s.match(/(\\d{1,2})[\\/\\-](\\d{1,2})[\\/\\-](\\d{4})/); return m?`${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`:s.slice(0,10);
}
function num(v){if(typeof v==="number")return v;return Number(String(v??"").replace(/R\\$\\s?/g,"").replace(/\\./g,"").replace(",","."))||0}
function rows(ws){return ws?XLSX.utils.sheet_to_json(ws,{defval:""}):[]}
function sheetFrom(arr){return XLSX.utils.json_to_sheet(arr||[])}

const defaults={
  lancamentos:["ID","Data","Tipo","Descrição","Categoria","Valor","Status","Observação"],
  contas:["ID","Descrição","Categoria","Valor","Vencimento","Status","Observação"],
  receitas:["ID","Descrição","Categoria","Valor","Data Prevista","Status","Observação"],
  categorias:["ID","Nome","Tipo"]
};

function emptyData(){
  state.lancamentos=[];state.contas=[];state.receitas=[];
  state.categorias=[
    {ID:uid(),Nome:"Alimentação",Tipo:"Despesa"},{ID:uid(),Nome:"Casa",Tipo:"Despesa"},
    {ID:uid(),Nome:"Energia",Tipo:"Despesa"},{ID:uid(),Nome:"Água",Tipo:"Despesa"},
    {ID:uid(),Nome:"Internet",Tipo:"Despesa"},{ID:uid(),Nome:"Transporte",Tipo:"Despesa"},
    {ID:uid(),Nome:"Salário",Tipo:"Receita"},{ID:uid(),Nome:"Vendas",Tipo:"Receita"},
    {ID:uid(),Nome:"Investimentos",Tipo:"Receita"},{ID:uid(),Nome:"Outros",Tipo:"Despesa"}
  ];
}
function showApp(){ $("#welcome").classList.add("hidden");$("#app").classList.remove("hidden");render(); }
async function createNew(){
  emptyData();
  state.workbook=null;
  state.fileHandle=null;
  state.fileName="KodsFin.xlsx";
  showApp();
  toast("Novo arquivo criado. Salve para escolher onde guardar o Excel.");
}

async function openWorkbookPicker(){
  if(!window.showOpenFilePicker){
    $("#fileInput").click();
    return;
  }
  try{
    const [handle]=await window.showOpenFilePicker({
      multiple:false,
      types:[{description:"Arquivo Excel",accept:{"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":[".xlsx"]}}],
      excludeAcceptAllOption:false
    });
    const file=await handle.getFile();
    state.fileHandle=handle;
    importWorkbook(file);
  }catch(err){
    if(err?.name!=="AbortError") console.error(err);
  }
}

function importWorkbook(file){
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:"array"});
      state.workbook=wb;state.fileName=file.name||"KodsFin.xlsx";
      state.lancamentos=rows(wb.Sheets["Lançamentos"]).map(x=>({...x,ID:x.ID||uid(),Data:normalizeDate(x.Data),Valor:num(x.Valor)}));
      state.contas=rows(wb.Sheets["Contas"]).map(x=>({...x,ID:x.ID||uid(),Vencimento:normalizeDate(x.Vencimento),Valor:num(x.Valor)}));
      state.receitas=rows(wb.Sheets["Receitas"]).map(x=>({...x,ID:x.ID||uid(),"Data Prevista":normalizeDate(x["Data Prevista"]),Valor:num(x.Valor)}));
      state.categorias=rows(wb.Sheets["Categorias"]).map(x=>({...x,ID:x.ID||uid()}));
      if(!state.categorias.length) emptyData();
      showApp(); toast(state.fileHandle?"Arquivo aberto para edição":"Arquivo carregado");
    }catch(err){alert("Não foi possível ler o Excel. Verifique se o arquivo é válido.");console.error(err)}
  };
  reader.readAsArrayBuffer(file);
}
function buildWorkbook(){
  const wb=XLSX.utils.book_new();
  const sets=[
    ["Lançamentos",state.lancamentos,defaults.lancamentos],
    ["Contas",state.contas,defaults.contas],
    ["Receitas",state.receitas,defaults.receitas],
    ["Categorias",state.categorias,defaults.categorias]
  ];
  sets.forEach(([name,data,headers])=>{
    const ws=sheetFrom(data.length?data:[Object.fromEntries(headers.map(h=>[h,""]))]);
    XLSX.utils.book_append_sheet(wb,ws,name);
  });
  return wb;
}
async function saveExcel(){
  const wb=buildWorkbook();
  const data=XLSX.write(wb,{bookType:"xlsx",type:"array"});

  try{
    // Se o arquivo foi aberto pelo seletor moderno, sobrescreve o mesmo arquivo.
    if(state.fileHandle){
      const permission=await state.fileHandle.queryPermission({mode:"readwrite"});
      if(permission!=="granted") await state.fileHandle.requestPermission({mode:"readwrite"});
      const writable=await state.fileHandle.createWritable();
      await writable.write(data);
      await writable.close();
      state.fileName=state.fileHandle.name;
      render();
      toast("Salvo no mesmo arquivo Excel");
      return;
    }

    // Se o arquivo foi carregado pelo input tradicional ou é um novo arquivo,
    // usa o seletor de salvamento para criar o arquivo e guardar a referência.
    if(window.showSaveFilePicker){
      const handle=await window.showSaveFilePicker({
        suggestedName:state.fileName||"KodsFin.xlsx",
        types:[{description:"Arquivo Excel",accept:{"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":[".xlsx"]}}]
      });
      const writable=await handle.createWritable();
      await writable.write(data);
      await writable.close();
      state.fileHandle=handle;
      state.fileName=handle.name;
      render();
      toast("Arquivo salvo. Os próximos salvamentos substituirão este arquivo.");
      return;
    }

    // Fallback para navegadores sem File System Access API.
    XLSX.writeFile(wb,state.fileName||"KodsFin.xlsx");
    toast("Excel baixado. Seu navegador não permite sobrescrever diretamente.");
  }catch(err){
    if(err?.name!=="AbortError"){
      console.error(err);
      toast("Não foi possível salvar o arquivo");
    }
  }
}

async function saveCopy(){
  const wb=buildWorkbook();
  try{
    if(window.showSaveFilePicker){
      const handle=await window.showSaveFilePicker({
        suggestedName:"KodsFin - Cópia.xlsx",
        types:[{description:"Arquivo Excel",accept:{"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":[".xlsx"]}}]
      });
      const data=XLSX.write(wb,{bookType:"xlsx",type:"array"});
      const writable=await handle.createWritable();
      await writable.write(data);
      await writable.close();
      toast("Cópia de segurança criada");
      return;
    }
    XLSX.writeFile(wb,"KodsFin - Cópia.xlsx");
  }catch(err){
    if(err?.name!=="AbortError") console.error(err);
  }
}
function badge(s){return `<span class="badge ${String(s||"").toLowerCase()}">${esc(s||"")}</span>`}
function dateBR(d){if(!d)return "-";const [y,m,day]=String(d).slice(0,10).split("-");return day?`${day}/${m}/${y}`:d}
function monthLabel(k){const [y,m]=k.split("-");return new Date(Number(y),Number(m)-1,1).toLocaleDateString("pt-BR",{month:"long",year:"numeric"})}
function currentMonth(){return todayISO().slice(0,7)}

function render(){
  $("#fileName").textContent=state.fileName;
  $("#todayLabel").textContent=new Date().toLocaleDateString("pt-BR",{weekday:"long",day:"2-digit",month:"long",year:"numeric"});
  document.querySelectorAll(".nav-btn").forEach(b=>b.classList.toggle("active",b.dataset.view===state.currentView));
  $("#pageTitle").textContent={dashboard:"Dashboard",lancamentos:"Lançamentos",contas:"Contas",receitas:"Receitas",planejamento:"Planejamento",categorias:"Categorias"}[state.currentView];
  ({dashboard:renderDashboard,lancamentos:renderLancamentos,contas:renderContas,receitas:renderReceitas,planejamento:renderPlanejamento,categorias:renderCategorias}[state.currentView])();
}
function renderDashboard(){
  const m=currentMonth();
  const rec=state.lancamentos.filter(x=>monthKey(x.Data)===m&&x.Tipo==="Receita"&&x.Status!=="Cancelado").reduce((a,x)=>a+num(x.Valor),0);
  const desp=state.lancamentos.filter(x=>monthKey(x.Data)===m&&x.Tipo==="Despesa"&&x.Status!=="Cancelado").reduce((a,x)=>a+num(x.Valor),0);
  const saldo=state.lancamentos.filter(x=>x.Status!=="Cancelado"&&(x.Status==="Pago"||x.Status==="Recebido")).reduce((a,x)=>a+(x.Tipo==="Receita"?num(x.Valor):-num(x.Valor)),0);
  const pendContas=state.contas.filter(x=>x.Status==="Pendente").reduce((a,x)=>a+num(x.Valor),0);
  const pendRec=state.receitas.filter(x=>x.Status==="Pendente").reduce((a,x)=>a+num(x.Valor),0);
  $("#content").innerHTML=`
    <div class="grid cards">
      <div class="card"><div class="card-label">Saldo</div><div class="card-value ${saldo>=0?"positive":"negative"}">${money(saldo)}</div></div>
      <div class="card"><div class="card-label">Receitas do mês</div><div class="card-value positive">${money(rec)}</div></div>
      <div class="card"><div class="card-label">Despesas do mês</div><div class="card-value negative">${money(desp)}</div></div>
      <div class="card"><div class="card-label">Resultado do mês</div><div class="card-value ${rec-desp>=0?"positive":"negative"}">${money(rec-desp)}</div></div>
    </div>
    <div class="grid two section">
      <div class="card chart-card"><div class="section-head"><h3>Receitas x Despesas</h3></div><div class="chart-wrap"><canvas id="monthlyChart"></canvas></div></div>
      <div class="card chart-card"><div class="section-head"><h3>Despesas por categoria</h3></div><div class="chart-wrap"><canvas id="categoryChart"></canvas></div></div>
    </div>
    <div class="grid two section">
      <div class="card"><div class="section-head"><h3>Contas pendentes</h3><strong>${money(pendContas)}</strong></div>${pendingRows(state.contas,"Vencimento")}</div>
      <div class="card"><div class="section-head"><h3>Receitas pendentes</h3><strong class="positive">${money(pendRec)}</strong></div>${pendingRows(state.receitas,"Data Prevista")}</div>
    </div>`;
  drawCharts();
}
function pendingRows(arr,dateField){
  const p=arr.filter(x=>x.Status==="Pendente").sort((a,b)=>String(a[dateField]).localeCompare(String(b[dateField]))).slice(0,5);
  return p.length?p.map(x=>`<div class="summary-row"><span>${esc(x.Descrição)}</span><span>${money(x.Valor)}</span></div>`).join(""):`<div class="empty">Nada pendente.</div>`;
}
function destroyCharts(){Object.values(state.charts).forEach(c=>c?.destroy());state.charts={}}
function drawCharts(){
  destroyCharts();
  if(typeof Chart==="undefined")return;
  const labels=[], rec=[], desp=[];
  const base=new Date(); base.setDate(1);
  for(let i=5;i>=0;i--){const d=new Date(base.getFullYear(),base.getMonth()-i,1),k=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;labels.push(d.toLocaleDateString("pt-BR",{month:"short"}));rec.push(state.lancamentos.filter(x=>monthKey(x.Data)===k&&x.Tipo==="Receita"&&x.Status!=="Cancelado").reduce((a,x)=>a+num(x.Valor),0));desp.push(state.lancamentos.filter(x=>monthKey(x.Data)===k&&x.Tipo==="Despesa"&&x.Status!=="Cancelado").reduce((a,x)=>a+num(x.Valor),0))}
  state.charts.month=new Chart($("#monthlyChart"),{type:"bar",data:{labels,datasets:[{label:"Receitas",data:rec},{label:"Despesas",data:desp}]},options:{responsive:true,maintainAspectRatio:false}});
  const cat={};state.lancamentos.filter(x=>monthKey(x.Data)===currentMonth()&&x.Tipo==="Despesa"&&x.Status!=="Cancelado").forEach(x=>cat[x.Categoria]=(cat[x.Categoria]||0)+num(x.Valor));
  state.charts.cat=new Chart($("#categoryChart"),{type:"doughnut",data:{labels:Object.keys(cat),datasets:[{data:Object.values(cat)}]},options:{responsive:true,maintainAspectRatio:false}});
}

function renderLancamentos(){
  $("#content").innerHTML=`
  <div class="toolbar">
    <input id="searchL" class="grow" placeholder="Buscar descrição...">
    <select id="filterType"><option value="">Todos</option><option>Receita</option><option>Despesa</option></select>
    <select id="filterStatus"><option value="">Todos os status</option><option>Pago</option><option>Recebido</option><option>Pendente</option><option>Cancelado</option></select>
    <button class="btn danger" onclick="openEntry('Despesa')">+ Despesa</button>
    <button class="btn success" onclick="openEntry('Receita')">+ Receita</button>
  </div><div id="launchTable" class="table-wrap"></div>`;
  ["searchL","filterType","filterStatus"].forEach(id=>$("#"+id).addEventListener("input",updateLaunchTable));
  updateLaunchTable();
}
function updateLaunchTable(){
  const q=($("#searchL")?.value||"").toLowerCase(),t=$("#filterType")?.value||"",s=$("#filterStatus")?.value||"";
  const data=state.lancamentos.filter(x=>(!q||String(x.Descrição).toLowerCase().includes(q))&&(!t||x.Tipo===t)&&(!s||x.Status===s)).sort((a,b)=>String(b.Data).localeCompare(String(a.Data)));
  $("#launchTable").innerHTML=data.length?`<table class="table"><thead><tr><th>Data</th><th>Tipo</th><th>Descrição</th><th>Categoria</th><th>Valor</th><th>Status</th><th>Ações</th></tr></thead><tbody>${data.map(x=>`<tr><td>${dateBR(x.Data)}</td><td>${esc(x.Tipo)}</td><td>${esc(x.Descrição)}</td><td>${esc(x.Categoria)}</td><td>${money(x.Valor)}</td><td>${badge(x.Status)}</td><td><button class="btn ghost" onclick="editEntry('${x.ID}')">Editar</button> <button class="btn ghost" onclick="deleteItem('lancamentos','${x.ID}')">Excluir</button></td></tr>`).join("")}</tbody></table>`:`<div class="empty">Nenhum lançamento encontrado.</div>`;
}

function renderContas(){
  $("#content").innerHTML=`<div class="toolbar"><input id="searchC" class="grow" placeholder="Buscar conta..."><select id="statusC"><option value="">Todos</option><option>Pendente</option><option>Pago</option><option>Cancelado</option></select><button class="btn primary" onclick="openAccount()">+ Nova conta</button></div><div id="contaTable" class="table-wrap"></div>`;
  $("#searchC").addEventListener("input",updateConta);$("#statusC").addEventListener("input",updateConta);updateConta();
}
function updateConta(){
  const q=($("#searchC")?.value||"").toLowerCase(),s=$("#statusC")?.value||"";
  const d=state.contas.filter(x=>(!q||String(x.Descrição).toLowerCase().includes(q))&&(!s||x.Status===s)).sort((a,b)=>String(a.Vencimento).localeCompare(String(b.Vencimento)));
  $("#contaTable").innerHTML=d.length?`<table class="table"><thead><tr><th>Vencimento</th><th>Descrição</th><th>Categoria</th><th>Valor</th><th>Status</th><th>Ações</th></tr></thead><tbody>${d.map(x=>`<tr><td>${dateBR(x.Vencimento)}</td><td>${esc(x.Descrição)}</td><td>${esc(x.Categoria)}</td><td>${money(x.Valor)}</td><td>${badge(x.Status)}</td><td>${x.Status==="Pendente"?`<button class="btn success" onclick="markPaid('${x.ID}')">Pagar</button>`:""} <button class="btn ghost" onclick="editAccount('${x.ID}')">Editar</button> <button class="btn ghost" onclick="deleteItem('contas','${x.ID}')">Excluir</button></td></tr>`).join("")}</tbody></table>`:`<div class="empty">Nenhuma conta encontrada.</div>`;
}
function renderReceitas(){
  $("#content").innerHTML=`<div class="toolbar"><input id="searchR" class="grow" placeholder="Buscar receita..."><select id="statusR"><option value="">Todos</option><option>Pendente</option><option>Recebido</option><option>Cancelado</option></select><button class="btn primary" onclick="openRevenue()">+ Nova receita</button></div><div id="recTable" class="table-wrap"></div>`;
  $("#searchR").addEventListener("input",updateRevenue);$("#statusR").addEventListener("input",updateRevenue);updateRevenue();
}
function updateRevenue(){
  const q=($("#searchR")?.value||"").toLowerCase(),s=$("#statusR")?.value||"";
  const d=state.receitas.filter(x=>(!q||String(x.Descrição).toLowerCase().includes(q))&&(!s||x.Status===s)).sort((a,b)=>String(a["Data Prevista"]).localeCompare(String(b["Data Prevista"])));
  $("#recTable").innerHTML=d.length?`<table class="table"><thead><tr><th>Data</th><th>Descrição</th><th>Categoria</th><th>Valor</th><th>Status</th><th>Ações</th></tr></thead><tbody>${d.map(x=>`<tr><td>${dateBR(x["Data Prevista"])}</td><td>${esc(x.Descrição)}</td><td>${esc(x.Categoria)}</td><td>${money(x.Valor)}</td><td>${badge(x.Status)}</td><td>${x.Status==="Pendente"?`<button class="btn success" onclick="markReceived('${x.ID}')">Receber</button>`:""} <button class="btn ghost" onclick="editRevenue('${x.ID}')">Editar</button> <button class="btn ghost" onclick="deleteItem('receitas','${x.ID}')">Excluir</button></td></tr>`).join("")}</tbody></table>`:`<div class="empty">Nenhuma receita encontrada.</div>`;
}
function renderPlanejamento(){
  const now=new Date(),end=new Date(now.getFullYear(),now.getMonth()+3,0),start=todayISO();
  const bills=state.contas.filter(x=>x.Status==="Pendente"&&x.Vencimento>=start&&x.Vencimento<=end.toISOString().slice(0,10)).reduce((a,x)=>a+num(x.Valor),0);
  const rev=state.receitas.filter(x=>x.Status==="Pendente"&&x["Data Prevista"]>=start&&x["Data Prevista"]<=end.toISOString().slice(0,10)).reduce((a,x)=>a+num(x.Valor),0);
  $("#content").innerHTML=`<div class="grid cards"><div class="card"><div class="card-label">Contas pendentes (próx. 3 meses)</div><div class="card-value negative">${money(bills)}</div></div><div class="card"><div class="card-label">Receitas previstas</div><div class="card-value positive">${money(rev)}</div></div><div class="card"><div class="card-label">Resultado previsto</div><div class="card-value ${rev-bills>=0?"positive":"negative"}">${money(rev-bills)}</div></div></div><div class="card section"><div class="section-head"><h3>Próximos compromissos</h3></div>${planningTable()}</div>`;
}
function planningTable(){
  const a=[...state.contas.filter(x=>x.Status==="Pendente").map(x=>({...x,_date:x.Vencimento,_type:"Conta"})),...state.receitas.filter(x=>x.Status==="Pendente").map(x=>({...x,_date:x["Data Prevista"],_type:"Receita"}))].filter(x=>x._date>=todayISO()).sort((a,b)=>a._date.localeCompare(b._date)).slice(0,20);
  return a.length?`<div class="table-wrap"><table class="table"><thead><tr><th>Data</th><th>Tipo</th><th>Descrição</th><th>Valor</th></tr></thead><tbody>${a.map(x=>`<tr><td>${dateBR(x._date)}</td><td>${x._type}</td><td>${esc(x.Descrição)}</td><td class="${x._type==="Receita"?"positive":"negative"}">${money(x.Valor)}</td></tr>`).join("")}</tbody></table></div>`:`<div class="empty">Nenhum compromisso futuro.</div>`;
}
function renderCategorias(){
  $("#content").innerHTML=`<div class="toolbar"><button class="btn primary" onclick="addCategory()">+ Nova categoria</button></div><div class="table-wrap"><table class="table"><thead><tr><th>Nome</th><th>Tipo</th><th>Ação</th></tr></thead><tbody>${state.categorias.map(x=>`<tr><td>${esc(x.Nome)}</td><td>${esc(x.Tipo)}</td><td><button class="btn ghost" onclick="deleteItem('categorias','${x.ID}')">Excluir</button></td></tr>`).join("")}</tbody></table></div>`;
}

function categoryOptions(type){
  return state.categorias.filter(x=>x.Tipo===type).map(x=>`<option value="${esc(x.Nome)}">${esc(x.Nome)}</option>`).join("");
}
function openEntry(type="Despesa",id=null){
  const old=state.lancamentos.find(x=>x.ID===id)||{ID:"",Data:todayISO(),Tipo:type,Descrição:"",Categoria:"",Valor:"",Status:type==="Despesa"?"Pago":"Recebido",Observação:""};
  $("#modalTitle").textContent=id?"Editar lançamento":`Nova ${type.toLowerCase()}`;
  $("#entryForm").innerHTML=`<div class="form-grid">
  <div class="form-group"><label>Data</label><input class="form-control" name="Data" type="date" value="${esc(old.Data)}" required></div>
  <div class="form-group"><label>Tipo</label><select class="form-control" name="Tipo">${["Receita","Despesa"].map(x=>`<option ${old.Tipo===x?"selected":""}>${x}</option>`).join("")}</select></div>
  <div class="form-group"><label>Descrição</label><input class="form-control" name="Descrição" value="${esc(old.Descrição)}" required></div>
  <div class="form-group"><label>Categoria</label><select class="form-control" name="Categoria"><option value="">Selecione</option>${categoryOptions(old.Tipo)}</select></div>
  <div class="form-group"><label>Valor</label><input class="form-control" name="Valor" type="number" step="0.01" min="0.01" value="${num(old.Valor)}" required></div>
  <div class="form-group"><label>Status</label><select class="form-control" name="Status">${["Pago","Recebido","Pendente","Cancelado"].map(x=>`<option ${old.Status===x?"selected":""}>${x}</option>`).join("")}</select></div>
  <div class="form-group full"><label>Observação</label><textarea class="form-control" name="Observação">${esc(old.Observação)}</textarea></div>
  </div><div class="form-actions"><button type="button" class="btn ghost" onclick="closeModal()">Cancelar</button><button class="btn primary">Salvar</button></div>`;
  $("#entryForm").querySelector('[name="Categoria"]').value=old.Categoria||"";
  $("#entryForm").querySelector('[name="Tipo"]').addEventListener("change",e=>{
    const sel=$("#entryForm").querySelector('[name="Categoria"]');sel.innerHTML='<option value="">Selecione</option>'+categoryOptions(e.target.value);
  });
  $("#entryForm").onsubmit=e=>{e.preventDefault();const fd=new FormData(e.target),o=Object.fromEntries(fd.entries());o.ID=old.ID||uid();o.Valor=num(o.Valor);if(id){const i=state.lancamentos.findIndex(x=>x.ID===id);state.lancamentos[i]=o}else state.lancamentos.push(o);closeModal();render();toast("Lançamento salvo")};
  $("#modal").classList.remove("hidden");
}
function editEntry(id){openEntry("",id)}
function openAccount(id=null){
  const old=state.contas.find(x=>x.ID===id)||{ID:"",Descrição:"",Categoria:"",Valor:"",Vencimento:todayISO(),Status:"Pendente",Observação:""};
  formModal("Nova conta",old,["Descrição","Categoria","Valor","Vencimento","Status","Observação"],o=>{
    o.ID=old.ID||uid();o.Valor=num(o.Valor); if(id)state.contas[state.contas.findIndex(x=>x.ID===id)]=o;else state.contas.push(o);closeModal();render();toast("Conta salva");
  },"conta");
}
function editAccount(id){openAccount(id)}
function openRevenue(id=null){
  const old=state.receitas.find(x=>x.ID===id)||{ID:"",Descrição:"",Categoria:"",Valor:"", "Data Prevista":todayISO(),Status:"Pendente",Observação:""};
  formModal("Nova receita",old,["Descrição","Categoria","Valor","Data Prevista","Status","Observação"],o=>{
    o.ID=old.ID||uid();o.Valor=num(o.Valor);if(id)state.receitas[state.receitas.findIndex(x=>x.ID===id)]=o;else state.receitas.push(o);closeModal();render();toast("Receita salva");
  },"receita");
}
function editRevenue(id){openRevenue(id)}
function formModal(title,old,fields,save,type){
  $("#modalTitle").textContent=title;
  const isConta=type==="conta", dateField=isConta?"Vencimento":"Data Prevista";
  $("#entryForm").innerHTML=`<div class="form-grid">
  <div class="form-group"><label>Descrição</label><input class="form-control" name="Descrição" value="${esc(old.Descrição)}" required></div>
  <div class="form-group"><label>Categoria</label><select class="form-control" name="Categoria"><option value="">Selecione</option>${categoryOptions("Despesa")}${type==="receita"?categoryOptions("Receita"):""}</select></div>
  <div class="form-group"><label>Valor</label><input class="form-control" name="Valor" type="number" step="0.01" min="0.01" value="${num(old.Valor)}" required></div>
  <div class="form-group"><label>${dateField}</label><input class="form-control" name="${dateField}" type="date" value="${esc(old[dateField])}" required></div>
  <div class="form-group"><label>Status</label><select class="form-control" name="Status">${(isConta?["Pendente","Pago","Cancelado"]:["Pendente","Recebido","Cancelado"]).map(x=>`<option ${old.Status===x?"selected":""}>${x}</option>`).join("")}</select></div>
  <div class="form-group full"><label>Observação</label><textarea class="form-control" name="Observação">${esc(old.Observação)}</textarea></div>
  </div><div class="form-actions"><button type="button" class="btn ghost" onclick="closeModal()">Cancelar</button><button class="btn primary">Salvar</button></div>`;
  $("#entryForm").querySelector('[name="Categoria"]').value=old.Categoria||"";
  $("#entryForm").onsubmit=e=>{e.preventDefault();save(Object.fromEntries(new FormData(e.target).entries()))};
  $("#modal").classList.remove("hidden");
}
function closeModal(){$("#modal").classList.add("hidden")}
function deleteItem(collection,id){if(!confirm("Excluir este registro?"))return;state[collection]=state[collection].filter(x=>x.ID!==id);render();toast("Registro excluído")}
function markPaid(id){const x=state.contas.find(x=>x.ID===id);if(x){x.Status="Pago";state.lancamentos.push({ID:uid(),Data:todayISO(),Tipo:"Despesa",Descrição:x.Descrição,Categoria:x.Categoria,Valor:num(x.Valor),Status:"Pago",Observação:"Gerado ao marcar conta como paga"});render();toast("Conta paga")}}
function markReceived(id){const x=state.receitas.find(x=>x.ID===id);if(x){x.Status="Recebido";state.lancamentos.push({ID:uid(),Data:todayISO(),Tipo:"Receita",Descrição:x.Descrição,Categoria:x.Categoria,Valor:num(x.Valor),Status:"Recebido",Observação:"Gerado ao marcar receita como recebida"});render();toast("Receita recebida")}}
function addCategory(){
  const name=prompt("Nome da categoria:");if(!name?.trim())return;
  const type=prompt("Tipo: Receita ou Despesa","Despesa");if(!["Receita","Despesa"].includes(type))return alert("Tipo inválido.");
  state.categorias.push({ID:uid(),Nome:name.trim(),Tipo:type});render();toast("Categoria adicionada");
}
function quick(type){openEntry(type)}
$("#fileInput").addEventListener("change",e=>e.target.files[0]&&importWorkbook(e.target.files[0]));
$("#newFileBtn").addEventListener("click",createNew);
$("#saveBtn").addEventListener("click",saveExcel);
$("#reloadBtn").addEventListener("click",()=>$("#fileInput").click());
$("#closeModal").addEventListener("click",closeModal);
$("#modal").addEventListener("click",e=>{if(e.target.id==="modal")closeModal()});
$("#quickExpense").addEventListener("click",()=>quick("Despesa"));
$("#quickIncome").addEventListener("click",()=>quick("Receita"));
$("#mobileMenu").addEventListener("click",()=>document.querySelector(".sidebar").classList.toggle("open"));
document.querySelectorAll(".nav-btn").forEach(b=>b.addEventListener("click",()=>{state.currentView=b.dataset.view;document.querySelector(".sidebar").classList.remove("open");render()}));
