// ---- Configuration: change these for your repo ----
const BASE = document.getElementById("gear-library").dataset.base;
// Primary source: a Google Sheet. Share it as "Anyone with the link can view", then paste its ID
// (the long string in the sheet URL between /d/ and /edit) and the name of the tab to read.
const SHEET_ID  = "1oEgN2sb9GXCfPA3CekSbCLOMXHhFhyfJ8UDudxoBPEI";                      // e.g. "1AbC...xyz"
const SHEET_TAB = "Gear";                  // tab name
// Optional: instead of the two lines above, use File > Share > Publish to web > CSV and paste that link here.
const SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQpI0C1iVCs4YuvRtEBJ_kIGRfEdMfB5aRvPxtHjFnYi5CKK9oZcpRodxP42O-T2EPEPoGVgxy_-Et0/pub?gid=0&single=true&output=csv";
// Backup source if the Google Sheet is unset or unreachable: an .xlsx file in this repo.
const XLSX_PATH = "gear.xlsx";
// Requests are submitted straight to a Google Form from the loadout panel. Set it up like this:
// 1. Create the form with these questions: name (short answer), email (short answer), pickup date and return date
//    (Date type), notes (paragraph), items (paragraph), total weight (short answer). Leave out any you don't want.
// 2. Form settings: do NOT require sign-in, do NOT limit to one response, no file uploads.
// 3. Three-dot menu > "Get pre-filled link": type any text in every question, click "Get link", copy it.
// 4. FORM_URL is that link up to the "?". Each entry.NNNNNNN in it is a question ID; match them to the keys below.
const FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSdLaaAPUUULcrH_PYgf27nW-2nnlNACTBCneg5qE-EfBOqf_Q/viewform";                       // e.g. "https://docs.google.com/forms/d/e/1FAIpQLS.../viewform"
const FORM_ENTRIES = {                     // fill in the IDs you have; blank ones are skipped
  name:   "entry.144521721",                              // e.g. "entry.1111111111"
  email:  "entry.245949958",
  pickup: "entry.2136646118",                              // Date question
  ret:    "entry.1132304307",                              // Date question (return date)
  notes:  "entry.2141839821",
  items:  "entry.2061083159"
};
const PH = "data:image/svg+xml," + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 80 60"><rect width="80" height="60" fill="#c5cfc9"/><path d="M0 50l22-26 14 16 10-12 34 22z" fill="#9fb0a8"/><circle cx="60" cy="16" r="6" fill="#e8edea"/></svg>');

// ---- Sample data (mirrors the spreadsheet columns); used when gear.xlsx can't be loaded ----
const COLS = ["ID","Item","Category","Subcategory","Type","Brand","Model","Fitment","Dimensions","Weight (g)","Condition","Status","Due back","Notes"];
const ROWS = [
 ["BAG-01","Handlebar roll 15L","Bags","Handlebar bag","Roll-top","Apidura","Expedition Handlebar Pack","Drop and flat bars, 31.8 mm clamp area","Ø 20 × 46 cm",420,"Good","Available","","Harness included"],
 ["BAG-02","Saddle pack 14L","Bags","Saddle bag","Roll-top","Revelate","Terrapin","Seatposts with 15+ cm exposed, no dropper posts","Ø 18 × 40 cm",510,"Like new","Checked out","2026-10-14","Dry bag liner included"],
 ["BAG-03","Frame bag, medium","Bags","Frame bag","Full frame","Ortlieb","Frame-Pack","Frames 54–58 cm with two bottle bosses","52 × 14 × 9 cm",360,"Fair","Available","","Roll-top, waterproof"],
 ["SLP-01","Ultralight tent, 1 person","Sleep system","Tent","Semi-freestanding","Big Agnes","Fly Creek HV UL1","Sleeps 1, fits a regular pad","Packed 41 × 15 cm",1050,"Good","Available","","Footprint included"],
 ["SLP-02","Bivy sack","Sleep system","Bivy","Waterproof","Outdoor Research","Helium Bivy","Fits regular pads up to 183 cm","213 × 76 cm",480,"Good","Available","",""],
 ["SLP-03","Down quilt, 20°F","Sleep system","Quilt","Down","Enlightened Equipment","Revelation","Fits sleepers up to 183 cm","Packed Ø 18 × 28 cm",640,"Like new","Checked out","2026-10-11",""],
 ["SLP-04","Inflatable pad, regular","Sleep system","Sleeping pad","Insulated","Therm-a-Rest","NeoAir XLite","Regular length, fits inside the bivy","183 × 51 × 6.3 cm",440,"Good","Available","","R-value 4.2"],
 ["KIT-01","Canister stove","Cook and water","Stove","Canister","MSR","PocketRocket 2","Threaded isobutane canisters","Packed 8 × 5 cm",73,"Good","Available","","Fuel not included"],
 ["KIT-02","Titanium pot, 750 ml","Cook and water","Cookware","Pot","Toaks","Titanium 750","Nests a 100 g fuel canister","Ø 9 × 12 cm",103,"Good","Available","",""],
 ["KIT-03","Gravity water filter, 2 L","Cook and water","Water filtration","Gravity","Platypus","GravityWorks 2L","Standard bottle threads","Packed 28 × 10 cm",320,"Fair","Available","",""],
 ["NAV-01","GPS head unit","Navigation and power","GPS","Head unit","Wahoo","ELEMNT Bolt","Out-front or stem mount, 31.8 mm bars","7.1 × 4.6 cm",99,"Good","Checked out","2026-10-20","Ohio maps preloaded"],
 ["NAV-02","Dynamo front light","Navigation and power","Lighting","Front light","Busch + Müller","IQ-X","Dynamo hub, fork eyelet mount","9 × 4 cm",150,"Good","Available","",""],
 ["REP-01","Trail repair kit","Repair","Tools","Kit","Library","Trail kit","Tubeless tires, 2–5 mm punctures","Roll 20 × 10 cm",260,"Like new","Available","","Multitool, plugs, quick link"],
 ["REP-02","Mini pump","Repair","Pumps","Hand pump","Lezyne","Pressure Drive","Presta and Schrader valves","Length 21 cm",130,"Good","Available","",""]
];
const SAMPLE = ROWS.map(r => Object.fromEntries(COLS.map((c,i)=>[c,r[i]])));

const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const pick = (r, ...keys) => { for (const k of keys) for (const rk in r) if (rk.trim().toLowerCase() === k) return r[rk]; return ""; };
const state = {items:[], q:"", cat:"All", sub:"", avail:false, sort:"item", loadout:new Set(), open:new Set(), source:""};

function fmtDate(v){
  if (typeof v === "number" && window.XLSX){ const d = XLSX.SSF.parse_date_code(v); if (d) return new Date(d.y,d.m-1,d.d).toLocaleDateString(undefined,{month:"short",day:"numeric"}); }
  const d = new Date(v); return isNaN(d) ? String(v||"") : d.toLocaleDateString(undefined,{month:"short",day:"numeric"});
}
function normalize(rows){
  return rows.map((r,n)=>{
    const id = String(pick(r,"id") || "ROW-"+n);
    let photos = String(pick(r,"photos","photo","images","image")).split(/[;,|]/).map(x=>x.trim()).filter(Boolean);
    if (!photos.length) photos = ["photos/"+id+".jpg"];
    return {
      id, photos,
      item: String(pick(r,"item","name")),
      cat: String(pick(r,"category") || "Other"),
      sub: String(pick(r,"subcategory")),
      type: String(pick(r,"type")),
      brand: String(pick(r,"brand")),
      model: String(pick(r,"model")),
      fit: String(pick(r,"fitment")),
      dims: String(pick(r,"dimensions","dimension")),
      weight: Number(pick(r,"weight (g)","weight")) || 0,
      cond: String(pick(r,"condition")),
      avail: /^avail/i.test(String(pick(r,"status"))),
      due: pick(r,"due back","due"),
      notes: String(pick(r,"notes"))
    };
  }).filter(x=>x.item);
}
async function fromSheet(){
  const url = SHEET_CSV_URL || (SHEET_ID && `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&headers=1&sheet=${encodeURIComponent(SHEET_TAB)}`);
  if(!url) throw new Error("no sheet configured");
  const res = await fetch(url);   // no custom fetch options: they can trigger a CORS preflight that Google rejects
  if(!res.ok) throw new Error("HTTP "+res.status);
  const text = await res.text();
  if(/^\s*<(!doctype|html)/i.test(text)) throw new Error("sheet is not shared publicly");
  const wb = XLSX.read(text, {type:"string", raw:true});
  const rows = normalize(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {raw:true, defval:""}));
  if(!rows.length) throw new Error("the sheet loaded but no rows have an Item column; check the headers in row 1 and the tab name");
  return rows;
}
async function fromXlsx(){
  const res = await fetch(XLSX_PATH, {cache:"no-cache"});
  if(!res.ok) throw new Error(res.status);
  const wb = XLSX.read(await res.arrayBuffer(), {type:"array"});
  return normalize(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]));
}
async function load(){
  try{ state.items = await fromSheet(); state.source = "sheet"; }
  catch(e1){
    state.err = /failed to fetch|networkerror|load failed/i.test(e1.message) ? "the browser could not reach it (wrong sheet ID, sheet not shared as Anyone with the link, or blocked by CORS)" : e1.message;
    try{ state.items = await fromXlsx(); state.source = "xlsx"; }
    catch(e2){ state.items = normalize(SAMPLE); state.source = "sample"; }
  }
  if(!state.items.length){ state.items = normalize(SAMPLE); state.source = "sample"; }
  buildFilters(); render();
}
function buildFilters(){
  const cats = ["All", ...[...new Set(state.items.map(i=>i.cat))].sort()];
  $("chips").innerHTML = cats.map(c=>`<button class="chip" aria-pressed="${c===state.cat}" data-cat="${esc(c)}">${esc(c)}</button>`).join("");
  const subs = [...new Set(state.items.filter(i=>state.cat==="All"||i.cat===state.cat).map(i=>i.sub).filter(Boolean))].sort();
  $("sub").innerHTML = `<option value="">All subcategories</option>` + subs.map(x=>`<option ${x===state.sub?"selected":""}>${esc(x)}</option>`).join("");
}
function visible(){
  const q = state.q.toLowerCase();
  const list = state.items.filter(i =>
    (state.cat==="All" || i.cat===state.cat) &&
    (!state.sub || i.sub===state.sub) &&
    (!state.avail || i.avail) &&
    (!q || [i.item,i.cat,i.sub,i.type,i.brand,i.model,i.fit,i.dims,i.notes].join(" ").toLowerCase().includes(q)));
  const s = state.sort;
  return list.sort((a,b)=> s==="weight" ? a.weight-b.weight : s==="category" ? a.cat.localeCompare(b.cat)||a.sub.localeCompare(b.sub)||a.item.localeCompare(b.item) : a.item.localeCompare(b.item));
}
function render(){
  const all = state.items, list = visible();
  $("stats").innerHTML = `<div><b>${all.length}</b>items in the library</div><div><b>${all.filter(i=>i.avail).length}</b>available now</div><div><b>${new Set(all.map(i=>i.cat)).size}</b>categories</div>`;
  $("source").innerHTML = state.source==="sheet" ? `<span class="dot"></span>Loaded live from Google Sheets`
    : state.source==="xlsx" ? `<span class="dot"></span>Loaded from ${esc(XLSX_PATH)} (Google Sheet problem: ${esc(state.err||"not set up")})`
    : `<span class="dot warn"></span>Showing sample data. ${state.err?`Google Sheet problem: ${esc(state.err)}.`:"Set SHEET_ID in the page script and share the sheet as Anyone with the link can view."} Photos show a placeholder until image files exist.`;
  const rows = list.map(i=>{
    const on = state.loadout.has(i.id), open = state.open.has(i.id);
    const spec = [["Type",i.type],["Brand",i.brand],["Model",i.model],["Fitment",i.fit],["Dimensions",i.dims],["Condition",i.cond],["Notes",i.notes]]
      .filter(x=>x[1]).map(x=>`<dt>${x[0]}</dt><dd>${esc(x[1])}</dd>`).join("");
    const more = i.photos.length>1 ? `<div class="more">${i.photos.map((p,n)=>`<button class="swap" data-swap="${esc(p)}" aria-label="Show photo ${n+1}"><img src="${esc(p)}" alt="" loading="lazy"></button>`).join("")}</div>` : "";
    return `<li class="item"><div class="row">
      <img class="thumb" src="${esc(i.photos[0])}" alt="" loading="lazy">
      <div><button class="name" data-open="${esc(i.id)}" aria-expanded="${open}">${esc(i.item)}</button><div class="sub">${esc([i.brand,i.model].filter(Boolean).join(" "))}</div></div>
      <div>${esc(i.cat)}<div class="sub">${esc(i.sub)}</div></div>
      <div class="num">${i.weight}<small> g</small></div>
      <div>${i.avail? '<span class="pill in">Available</span>' : `<span class="pill out">Out</span>${i.due? `<div class="sub">Back ${esc(fmtDate(i.due))}</div>`:""}`}</div>
      <div><button class="btn ${on?"on":""}" data-id="${esc(i.id)}" ${!i.avail&&!on?"disabled":""}>${on?"Remove":i.avail?"Add to loadout":"Unavailable"}</button></div>
    </div>
    <div class="detail" ${open?"":"hidden"}>
      <div class="photos"><img class="main" src="${esc(i.photos[0])}" alt="${esc(i.item)}">${more}</div>
      <dl class="specs">${spec}</dl>
    </div></li>`}).join("");
  const head = $("rows").querySelector(".head").outerHTML;
  $("rows").innerHTML = head + (rows || `<li class="empty">No gear matches those filters. Clear the search or switch category.</li>`);
  renderLoadout();
}
function renderLoadout(){
  const picked = state.items.filter(i=>state.loadout.has(i.id));
  $("count").textContent = picked.length ? `${picked.length} item${picked.length>1?"s":""}` : "Nothing added yet";
  $("tray").innerHTML = picked.map(i=>`<li><span>${esc(i.item)}</span><button data-rm="${esc(i.id)}" aria-label="Remove ${esc(i.item)}">✕</button></li>`).join("");
  $("send").disabled = !picked.length;
}
const flip = (set,id) => set.has(id) ? set.delete(id) : set.add(id);

$("q").addEventListener("input",e=>{state.q=e.target.value;render()});
$("sort").addEventListener("change",e=>{state.sort=e.target.value;render()});
$("sub").addEventListener("change",e=>{state.sub=e.target.value;render()});
$("avail").addEventListener("change",e=>{state.avail=e.target.checked;render()});
$("chips").addEventListener("click",e=>{const b=e.target.closest("[data-cat]"); if(!b)return; state.cat=b.dataset.cat; state.sub=""; buildFilters(); render();});
$("rows").addEventListener("click",e=>{
  const add=e.target.closest("[data-id]"), op=e.target.closest("[data-open]"), sw=e.target.closest("[data-swap]");
  if(add&&!add.disabled){ flip(state.loadout,add.dataset.id); render(); }
  else if(op){ flip(state.open,op.dataset.open); render(); }
  else if(sw){ const m=sw.closest(".photos").querySelector(".main"); m.removeAttribute("data-fb"); m.src=sw.dataset.swap; }
});
$("tray").addEventListener("click",e=>{const b=e.target.closest("[data-rm]"); if(b){ flip(state.loadout,b.dataset.rm); render(); }});
const today = new Date(Date.now()-new Date().getTimezoneOffset()*60000).toISOString().slice(0,10);
$("f-pickup").min = today; $("f-return").min = today;
$("f-pickup").addEventListener("change",e=>{ $("f-return").min = e.target.value || today; });

function formPayload(picked, iso){
  const E = FORM_ENTRIES, d = new URLSearchParams();
  const put = (k,v)=>{ if(E[k] && v) d.set(E[k], v); };
  const date = (k,v)=>{ if(!E[k]||!v) return; if(iso) d.set(E[k],v); else { const [y,m,dd]=v.split("-"); d.set(E[k]+"_year",y); d.set(E[k]+"_month",+m); d.set(E[k]+"_day",+dd); } };
  put("name",$("f-name").value.trim()); put("email",$("f-email").value.trim());
  date("pickup",$("f-pickup").value); date("ret",$("f-return").value);
  put("notes",$("f-notes").value.trim());
  put("items",picked.map(i=>`${i.id}: ${i.item}`).join("\n"));
  return d;
}
$("req").addEventListener("submit", async e=>{
  e.preventDefault();
  const msg = $("formmsg"), btn = $("send");
  const picked = state.items.filter(i=>state.loadout.has(i.id));
  if(!FORM_URL){ msg.textContent = "Requests are switched off until FORM_URL is set in the page script."; return; }
  if(!picked.length){ msg.textContent = "Add at least one item first."; return; }
  if($("f-return").value < $("f-pickup").value){ msg.textContent = "Return date must be on or after the pickup date."; return; }
  btn.disabled = true; msg.textContent = "Sending…";
  try{
    // A static page can't read Google's reply (no-cors), so "sent" means the request left the browser.
    await fetch(FORM_URL.replace(/\/viewform.*$/,"/formResponse"), {method:"POST", mode:"no-cors", body:formPayload(picked,false)});
    msg.textContent = "Request sent. We'll email you to confirm.";
    state.loadout.clear(); $("req").reset(); render();
  }catch(err){
    const link = FORM_URL + "?usp=pp_url&" + formPayload(picked,true).toString();
    msg.innerHTML = `Couldn't send. <a href="${esc(link)}" target="_blank" rel="noopener">Open the form instead</a>`;
    btn.disabled = false;
  }
});
// Any image that fails to load (missing file, bad URL) gets the placeholder
document.addEventListener("error",e=>{ const t=e.target; if(t.tagName==="IMG"&&!t.dataset.fb){ t.dataset.fb="1"; t.src=PH; } },true);
load();
