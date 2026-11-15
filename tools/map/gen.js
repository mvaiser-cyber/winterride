const topo=require('topojson-client'),d3=require('d3-geo'),fs=require('fs');
const us=JSON.parse(fs.readFileSync(__dirname+'/node_modules/us-atlas/states-10m.json'));
const states=topo.feature(us,us.objects.states).features;
const KEEP={'12':1,'13':1,'45':1,'37':1,'01':1,'47':1,'51':1},TRIP=['12','13','45','37'];
const W=440,H=700;
const fit=(ext,x0,y0,x1,y1)=>d3.geoMercator().fitExtent([[x0,y0],[x1,y1]],{type:'Feature',geometry:{type:'MultiPoint',coordinates:ext}});
const proj=fit([[-85.0,24.9],[-77.4,36.25]],0,0,W,H),path=d3.geoPath(proj).digits(1);
const P=(lon,lat,pr=proj)=>pr([lon,lat]).map(v=>+v.toFixed(1));
const C={amelia:'#1b86a4',aiken:'#d9732f',cabin:'#2c8a5a',chs:'#7457c2',sav:'#cc4a80',orl:'#2f68d0',home:'#203047',north:'#176c83',south:'#e0803a',air:'#1d2b3a'};
const stops=[
 {k:'home',n:'🏠',lon:-80.40,lat:26.10,name:'Weston',sub:'Start & finish',dx:16,dy:0},
 {k:'amelia',n:1,lon:-81.46,lat:30.67,name:'Amelia Island',sub:'Dec 22 · 1 night',dx:16,dy:0},
 {k:'aiken',n:2,lon:-81.72,lat:33.56,name:'Aiken',sub:'Dec 23 · 1 night',dx:-16,dy:0,anchor:'end'},
 {k:'cabin',n:3,lon:-82.40,lat:35.52,name:'Mountain cabin',sub:'Dec 24–30 · 6 nights',sub2:'✈ AVL airport ~20–25 min',dx:24,dy:-6},
 {k:'chs',n:4,lon:-79.93,lat:32.78,name:'Charleston',sub:'Dec 30 · 1 night',dx:16,dy:0},
 {k:'sav',n:5,lon:-81.09,lat:32.08,name:'Savannah',sub:'Dec 31 · 1 night',dx:16,dy:4},
 {k:'orl',n:6,lon:-81.38,lat:28.54,name:'Orlando',sub:'Jan 1–3 · 2 nights',dx:16,dy:0}];
const by=k=>stops.find(s=>s.k===k);
const legs=[['home','amelia','5.5–6 h','north',.75],['amelia','aiken','4.5–5.5 h','north'],['aiken','cabin','3–4 h','north'],['cabin','chs','4–4.5 h','south'],['chs','sav','2–2.5 h','south',.5,0,[-79.35,31.7]],['sav','orl','4.5–5 h','south',.62],['orl','home','3–3.5 h','south']];
const esc=t=>String(t).replace(/&/g,'&amp;').replace(/</g,'&lt;');
const txt=(x,y,t,o={})=>`<text x="${x}" y="${y}" font-size="${o.size||11}" font-weight="${o.w||700}" fill="${o.fill||'#203047'}" text-anchor="${o.anchor||'start'}"${o.style?` font-style="${o.style}"`:''}${o.ls?` letter-spacing="${o.ls}"`:''}${o.halo?' paint-order="stroke" stroke="#fff" stroke-width="3.5" stroke-linejoin="round"':''}>${esc(t)}</text>`;
let s=`<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-labelledby="mapT mapD" font-family="Inter,ui-sans-serif,system-ui,-apple-system,Segoe UI,Arial">`;
s+=`<title id="mapT">Road trip map with numbered overnight stops</title><desc id="mapD">Weston, then overnight stops: 1 Amelia Island, 2 Aiken, 3 the mountain cabin near Asheville, 4 Charleston, 5 Savannah, 6 Orlando, and back to Weston. Asheville airport (AVL) is about 20 minutes from the cabin area.</desc>`;
s+=`<defs><marker id="ahn" viewBox="0 0 10 10" refX="6" refY="5" markerUnits="userSpaceOnUse" markerWidth="11" markerHeight="11" orient="auto"><path d="M0 0L10 5L0 10z" fill="${C.north}"/></marker><marker id="ahs" viewBox="0 0 10 10" refX="6" refY="5" markerUnits="userSpaceOnUse" markerWidth="11" markerHeight="11" orient="auto"><path d="M0 0L10 5L0 10z" fill="${C.south}"/></marker><filter id="sh" x="-30%" y="-30%" width="160%" height="160%"><feDropShadow dx="0" dy="1" stdDeviation="1.3" flood-color="#0b2236" flood-opacity=".28"/></filter></defs>`;
s+=`<rect width="${W}" height="${H}" fill="#d7eaf4"/>`;
for(const f of states){if(!KEEP[f.id])continue;const t=TRIP.includes(f.id);s+=`<path d="${path(f)}" fill="${t?'#fbf7ee':'#ebe7dd'}" stroke="#c4bba6" stroke-width="${t?1:.7}" stroke-linejoin="round"/>`}
const lab=(t,lon,lat)=>{const[x,y]=P(lon,lat);return txt(x,y,t,{size:10.5,w:800,fill:'#bdb39c',anchor:'middle',ls:2})};
s+=lab('FLORIDA',-82.7,28.75)+lab('GEORGIA',-83.6,32.3)+lab('SOUTH',-79.75,34.35)+lab('CAROLINA',-79.75,34.15)+lab('NORTH CAROLINA',-79.0,35.85);
{const[x,y]=P(-78.9,29.4);s+=txt(x,y,'ATLANTIC',{size:11,w:600,fill:'#86abc0',anchor:'middle',style:'italic',ls:2.5})+txt(x,y+15,'OCEAN',{size:11,w:600,fill:'#86abc0',anchor:'middle',style:'italic',ls:2.5})}
// route legs (gently curved; northbound bows west, southbound bows east)
const pills=[];
for(const [a,b,t,dir,pt=.5,po=.95,fix] of legs){const A=P(by(a).lon,by(a).lat),B=P(by(b).lon,by(b).lat);
 const dx=B[0]-A[0],dy=B[1]-A[1],L=Math.hypot(dx,dy),sg=1,nx=dy/L*sg*16,ny=-dx/L*sg*16;
 const cx=(A[0]+B[0])/2+nx,cy=(A[1]+B[1])/2+ny;
 const q=(t)=>[(1-t)*(1-t)*A[0]+2*(1-t)*t*cx+t*t*B[0],(1-t)*(1-t)*A[1]+2*(1-t)*t*cy+t*t*B[1]];
 s+=`<path d="M${A} Q${cx.toFixed(1)},${cy.toFixed(1)} ${B}" fill="none" stroke="${C[dir]}" stroke-width="3.2" stroke-linecap="round"/>`;
 const m1=q(.47),m2=q(.53);s+=`<path d="M${m1.map(v=>v.toFixed(1))} L${m2.map(v=>v.toFixed(1))}" stroke="${C[dir]}" stroke-width="3.2" marker-end="url(#ah${dir[0]})"/>`;
 const m=q(pt);if(fix){const[fx,fy]=P(...fix);pills.push({x:fx,y:fy,t,dir})}else pills.push({x:m[0]+nx*po,y:m[1]+ny*po,t,dir})}
for(const p of pills){const w=p.t.length*5.5+14;s+=`<g filter="url(#sh)"><rect x="${(p.x-w/2).toFixed(1)}" y="${(p.y-8).toFixed(1)}" width="${w.toFixed(1)}" height="16" rx="8" fill="#fff"/></g>`+txt(+(p.x).toFixed(1),+(p.y+3.5).toFixed(1),p.t,{size:9.5,w:800,fill:C[p.dir],anchor:'middle'})}
const AVL=[-82.54,35.436];
// stops
for(const st of stops){const[x,y]=P(st.lon,st.lat);
 s+=`<g filter="url(#sh)"><circle cx="${x}" cy="${y}" r="12" fill="${C[st.k]}" stroke="#fff" stroke-width="2.5"/></g>`;
 s+=st.k==='home'?txt(x,y+4,'⌂',{size:15,w:900,fill:'#fff',anchor:'middle'}):txt(x,y+4.3,st.n,{size:12.5,w:900,fill:'#fff',anchor:'middle'});
 const ax=x+st.dx,ay=y+st.dy;s+=txt(ax,ay-1,st.name,{size:12,w:800,anchor:st.anchor,halo:1})+txt(ax,ay+12,st.sub,{size:10,w:600,fill:'#55677a',anchor:st.anchor,halo:1})+(st.sub2?txt(ax,ay+25,st.sub2,{size:10,w:800,fill:C.air,anchor:st.anchor,halo:1}):'')}
// inset: cabin area close-up (top-left)
const IX=10,IW=186,IH=176,IY=H-IH-10;
const ip=fit([[-82.86,35.16],[-82.16,35.70]],IX+6,IY+24,IX+IW-6,IY+IH-24);const ipath=d3.geoPath(ip).digits(1);
s+=`<g filter="url(#sh)"><rect x="${IX}" y="${IY}" width="${IW}" height="${IH}" rx="12" fill="#fff"/></g><clipPath id="ic"><rect x="${IX}" y="${IY}" width="${IW}" height="${IH}" rx="12"/></clipPath><g clip-path="url(#ic)"><rect x="${IX}" y="${IY+20}" width="${IW}" height="${IH-20}" fill="#f3f7ef"/>`;
for(const f of states){if(f.id==='37'||f.id==='45')s+=`<path d="${ipath(f)}" fill="${f.id==='37'?'#f3f7ef':'#ebe7dd'}" stroke="#c4bba6" stroke-width="1"/>`}
const ipt=(lon,lat)=>ip([lon,lat]).map(v=>+v.toFixed(1));
const towns=[['Asheville',-82.551,35.595,'end',-6,-5],['Black Mtn',-82.321,35.618,'start',6,-5],['Hendersonville',-82.461,35.318,'start',6,4],['Brevard',-82.734,35.233,'start',6,4]];
const cab=ipt(-82.40,35.52),air=ipt(...AVL);
// drive lines
const dl=(to,label,lx,ly)=>`<path d="M${air} L${to}" stroke="${C.air}" stroke-width="1.6" stroke-dasharray="4 3" fill="none"/>`+txt(lx,ly,label,{size:9.5,w:800,fill:C.air,anchor:'middle',halo:1});
const hv=ipt(-82.461,35.318);
s+=dl(cab,'~20–25 min',air[0]-22,air[1]-19)+dl(hv,'~15–20 min',(air[0]+hv[0])/2+34,(air[1]+hv[1])/2+2);
for(const [n,lo,la,an,dx,dy] of towns){const[x,y]=ipt(lo,la);s+=`<circle cx="${x}" cy="${y}" r="2.6" fill="#8796a5"/>`+txt(x+dx,y+dy,n,{size:9,w:700,fill:'#55677a',anchor:an,halo:1})}
s+=`<circle cx="${air[0]}" cy="${air[1]}" r="9" fill="${C.air}" stroke="#fff" stroke-width="2"/>`+txt(air[0],air[1]+3.6,'✈',{size:10,w:700,fill:'#fff',anchor:'middle'})+txt(air[0]-12,air[1]+3,'AVL',{size:10,w:900,fill:C.air,anchor:'end',halo:1});
s+=`<circle cx="${cab[0]}" cy="${cab[1]}" r="9.5" fill="${C.cabin}" stroke="#fff" stroke-width="2"/>`+txt(cab[0],cab[1]+3.8,'3',{size:11,w:900,fill:'#fff',anchor:'middle'})+txt(cab[0]+13,cab[1]+4,'Fairview',{size:9,w:800,fill:'#2c8a5a',halo:1});
s+=`</g><rect x="${IX}" y="${IY}" width="${IW}" height="20" rx="12" fill="${C.cabin}"/><rect x="${IX}" y="${IY+10}" width="${IW}" height="10" fill="${C.cabin}"/>`+txt(IX+10,IY+14,'Cabin area close-up',{size:10.5,w:800,fill:'#fff'});
s+=txt(IX+IW/2,IY+IH-7,'3 = cabin search area (exact spot TBD)',{size:8.5,w:600,fill:'#55677a',anchor:'middle'});
{const[x,y]=P(-82.40,35.52);s+=`<rect x="${x-19}" y="${y-19}" width="38" height="38" rx="6" fill="none" stroke="${C.cabin}" stroke-width="1.3" stroke-dasharray="3 3"/>`}
s+='</svg>';
fs.writeFileSync(process.argv[2],s);console.log('bytes',s.length);
