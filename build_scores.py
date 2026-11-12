import openpyxl, json, os
from datetime import datetime, timedelta

BASE = r'D:\DATA Drive D\ควบคุมคุณภาพ\Action Plan\2026\Engagement'
BAD  = {0x201C,0x201D,0x2018,0x2019,0x2013,0x2014,0x0060,0x00AB,0x00BB}

# ── 1. read Excel ──────────────────────────────────────────
wb = openpyxl.load_workbook(os.path.join(BASE,'สะสมคะแนน.xlsx'), data_only=True)
ws = wb.active
rows = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i==0: continue
    dv,name,menu = row[0],row[1],row[2]
    hc,fd,lb,ef,et,cl = row[3],row[4],row[5],row[6],row[7],row[9]
    if not dv or not name or not isinstance(dv,datetime): continue
    clean = ''.join(c for c in (menu or '') if ord(c) not in BAD)
    rows.append({
        'date':       dv.strftime('%Y-%m-%d'),
        'name':       name.strip(),
        'menu':       clean.strip(),
        'homeChef':   int(hc) if hc else 0,
        'foodie':     int(fd) if fd else 0,
        'lunchbox':   int(lb) if lb else 0,
        'extraFriend':int(ef) if ef else 0,
        'extraTask':  int(et) if isinstance(et,(int,float)) else 0,
        'countLike':  int(cl) if isinstance(cl,(int,float)) else 0,
    })

print(f'rows={len(rows)}')

# verify no bad chars
bad_found = [c for r in rows for c in r['menu'] if ord(c) in BAD]
print(f'bad chars: {len(bad_found)}')

# ── 2. build WEEKS ─────────────────────────────────────────
dates = sorted({r['date'] for r in rows})
dt_min = datetime.strptime(dates[0],'%Y-%m-%d')
dt_max = datetime.strptime(dates[-1],'%Y-%m-%d')
weeks=[]; wk=1; start=dt_min
while start<=dt_max:
    end=min(start+timedelta(days=6),dt_max)
    weeks.append({'label':f'Week {wk}','from':start.strftime('%Y-%m-%d'),'to':end.strftime('%Y-%m-%d')})
    wk+=1; start=end+timedelta(days=1)
weeks_all=[{'label':'รวมทั้งหมด','from':dates[0],'to':dates[-1]}]+weeks

# ── 3. embed data as JSON (ensure_ascii=True = safe) ───────
raw_js   = 'const RAW='   + json.dumps(rows,      ensure_ascii=True) + ';'
weeks_js = 'const WEEKS=' + json.dumps(weeks_all, ensure_ascii=True) + ';'
ts = datetime.now().strftime('%d/%m/%Y %H:%M')

# ── 4. HTML ────────────────────────────────────────────────
html = """<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>""" + f'\u0e2a\u0e23\u0e38\u0e1b\u0e04\u0e30\u0e41\u0e19\u0e19\u0e2a\u0e30\u0e2a\u0e21 \u0e08\u0e32\u0e19\u0e40\u0e14\u0e47\u0e14\u0e40\u0e2a\u0e23\u0e47\u0e08\u0e41\u0e19\u0e48 <!-- {ts} -->' + """</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI','Sarabun',sans-serif;background:#f5f0eb;min-height:100vh;padding:20px 12px 48px;color:#2d2018}
h1{text-align:center;font-size:clamp(1.3rem,4vw,2rem);color:#8b3a1a;padding:28px 20px 20px}
.box{background:#fff;border:1.5px solid #e8d8c8;border-radius:14px;padding:16px 20px;max-width:720px;margin:0 auto 20px;box-shadow:0 2px 10px rgba(180,120,60,.08)}
.row{display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end}
label{font-size:.8rem;color:#8b6040;font-weight:600;display:block;margin-bottom:4px}
input{border:1.5px solid #ddd0c0;border-radius:8px;padding:8px 12px;font-size:.95rem;width:140px;background:#faf7f3;color:#2d2018;outline:none}
input:focus{border-color:#c8722a}
.btn{background:#c8722a;color:#fff;border:none;border-radius:8px;padding:9px 22px;font-size:.9rem;font-weight:700;cursor:pointer}
.btn:hover{background:#a85a1e}
.hint{font-size:.7rem;color:#bbb;margin-top:3px}
.msg{font-size:.85rem;color:#e55}
.tabs{display:flex;flex-wrap:wrap;gap:8px;justify-content:center;margin-bottom:20px}
.tab{padding:6px 16px;border-radius:999px;border:2px solid #c8722a;background:#fff;color:#c8722a;cursor:pointer;font-size:.82rem;font-weight:600}
.tab:hover,.tab.on{background:#c8722a;color:#fff}
.tab.all{border-color:#a08060;color:#a08060}
.tab.all:hover,.tab.all.on{background:#a08060;color:#fff}
.podium{display:flex;justify-content:center;align-items:flex-end;gap:12px;margin-bottom:20px;flex-wrap:wrap}
.pc{text-align:center;border-radius:14px;padding:12px 16px;background:#fff;border:1.5px solid #e8d8c8;min-width:100px;box-shadow:0 2px 10px rgba(180,120,60,.08)}
.pc.p1{border-color:#c8722a;background:linear-gradient(135deg,#fff8f3,#fdeede);transform:translateY(-12px)}
.pc.p2{border-color:#b0a090}.pc.p3{border-color:#cd7f32}
.pm{font-size:1.8rem}.pn{font-weight:700;color:#2d2018;font-size:.95rem;margin-top:4px}
.pp{font-size:1.2rem;font-weight:800;color:#c8722a}
.cards{display:flex;flex-wrap:wrap;gap:10px;justify-content:center;margin-bottom:20px}
.card{background:#fff;border:1.5px solid #e8d8c8;border-radius:12px;padding:10px 18px;text-align:center;min-width:120px}
.cv{font-size:1.3rem;font-weight:800;color:#c8722a}
.cl{font-size:.72rem;color:#a08060;margin-top:2px}
.sec{max-width:1000px;margin:0 auto 24px}
.th{background:linear-gradient(135deg,#fff8f3,#fdeede);padding:12px 18px;display:flex;justify-content:space-between;align-items:center;border-radius:12px 12px 0 0;border-bottom:2px solid rgba(200,114,42,.2)}
.th-t{font-weight:700;color:#8b3a1a;font-size:.92rem}
.th-s{font-size:.8rem;color:#a08060}
.tw{overflow-x:auto;border-radius:0 0 12px 12px;border:1.5px solid #e8d8c8;border-top:none}
table{width:100%;border-collapse:collapse;background:#fff;font-size:.85rem}
thead tr{background:#fdf5ee}
th{padding:10px 12px;text-align:center;color:#8b3a1a;font-weight:700;border-bottom:2px solid rgba(200,114,42,.15);white-space:nowrap}
th:first-child{text-align:left;padding-left:16px}
tbody tr{border-bottom:1px solid #f0e8dc}
tbody tr:hover{background:#fdf8f3}
td{padding:9px 12px;text-align:center;color:#4a3020}
td:first-child{text-align:left;padding-left:16px;font-weight:600;color:#2d2018;white-space:nowrap}
.rb{display:inline-block;width:22px;height:22px;line-height:22px;border-radius:50%;font-size:.7rem;font-weight:700;margin-right:6px;text-align:center}
.r1{background:#c8722a;color:#fff}.r2{background:#b0a090;color:#fff}.r3{background:#cd7f32;color:#fff}.rn{background:#e8d8c8;color:#8b6040}
.pts{font-weight:800;color:#c8722a}.lk{color:#a08060;font-weight:600}.z{color:#ccc}
.ck{cursor:pointer;text-decoration:underline dotted}
.ck:hover{color:#c8722a}
.fbar{padding:10px 14px;background:#faf7f3;border-bottom:1px solid #f0e8dc;display:flex;flex-wrap:wrap;gap:7px;align-items:center}
.fl{font-size:.76rem;color:#a08060;font-weight:600;margin-right:4px}
.np{padding:3px 12px;border-radius:999px;border:1.5px solid #ddd0c0;background:#fff;color:#8b6040;cursor:pointer;font-size:.78rem}
.np:hover{border-color:#c8722a;color:#c8722a}
.np.on{background:#c8722a;border-color:#c8722a;color:#fff;font-weight:700}
.np.all{border-color:#a08060;color:#a08060}
.np.all.on{background:#a08060;color:#fff}
.td{color:#5b7fa6;white-space:nowrap;font-size:.78rem;font-weight:600}
.tn{color:#8b3a1a;font-weight:700;white-space:nowrap}
.tag{display:inline-block;padding:1px 7px;border-radius:999px;font-size:.68rem;font-weight:700;margin:1px}
.hc{background:#e8f5e9;color:#2e7d32}.fd{background:#fff3e0;color:#e65100}
.lb{background:#fff8e1;color:#f57f17}.et{background:#f3e5f5;color:#7b1fa2}.ef{background:#fce4ec;color:#c2185b}
.nd{text-align:center;padding:28px;color:#bbb}
.mbg{display:none;position:fixed;inset:0;background:rgba(100,60,20,.4);z-index:100;justify-content:center;align-items:center;padding:16px}
.mbg.open{display:flex}
.modal{background:#fff;border:1.5px solid #e8d8c8;border-radius:16px;max-width:560px;width:100%;max-height:80vh;overflow-y:auto;padding:22px;box-shadow:0 8px 40px rgba(180,120,60,.15)}
.modal h2{color:#8b3a1a;margin-bottom:14px;font-size:1rem}
.mc{float:right;background:none;border:none;color:#a08060;font-size:1.3rem;cursor:pointer}
.dr{display:flex;gap:6px;align-items:center;padding:5px 0;border-bottom:1px solid #f0e8dc;flex-wrap:wrap}
.dd{color:#5b7fa6;font-size:.76rem;min-width:70px;font-weight:600}
.dm{color:#2d2018;font-size:.83rem;flex:1}
.dl{color:#c8722a;font-size:.76rem;font-weight:600}
@media(max-width:600px){input{width:100%}table{font-size:.78rem}th,td{padding:7px 8px}}
</style></head>
<body>
<h1>&#x1F37D; &#x0E2A;&#x0E23;&#x0E38;&#x0E1B;&#x0E04;&#x0E30;&#x0E41;&#x0E19;&#x0E19;&#x0E2A;&#x0E30;&#x0E2A;&#x0E21; &quot;&#x0E08;&#x0E32;&#x0E19;&#x0E40;&#x0E14;&#x0E47;&#x0E14;&#x0E40;&#x0E2A;&#x0E23;&#x0E47;&#x0E08;&#x0E41;&#x0E19;&#x0E48;&quot;</h1>
<div class="box"><div class="row">
  <div><label>&#x0E27;&#x0E31;&#x0E19;&#x0E17;&#x0E35;&#x0E48;&#x0E40;&#x0E23;&#x0E34;&#x0E48;&#x0E21;&#x0E15;&#x0E49;&#x0E19;</label>
  <input id="f" placeholder="09/10/2026" maxlength="10" oninput="sl(this)"/>
  <div class="hint">&#x0E27;&#x0E27;/&#x0E14;&#x0E14;/&#x0E1B;&#x0E1B;&#x0E1B;&#x0E1B;</div></div>
  <div><label>&#x0E27;&#x0E31;&#x0E19;&#x0E17;&#x0E35;&#x0E48;&#x0E2A;&#x0E34;&#x0E49;&#x0E19;&#x0E2A;&#x0E38;&#x0E14;</label>
  <input id="t" placeholder="15/10/2026" maxlength="10" oninput="sl(this)"/>
  <div class="hint">&#x0E27;&#x0E27;/&#x0E14;&#x0E14;/&#x0E1B;&#x0E1B;&#x0E1B;&#x0E1B;</div></div>
  <button class="btn" onclick="go()">&#x0E04;&#x0E33;&#x0E19;&#x0E27;&#x0E13;</button>
  <span class="msg" id="mg"></span>
</div></div>
<div class="tabs" id="tabs"></div>
<div class="podium" id="pod"></div>
<div class="cards" id="cds"></div>
<div class="sec"><div class="th"><span class="th-t">&#x1F4CA; &#x0E1C;&#x0E25;&#x0E04;&#x0E30;&#x0E41;&#x0E19;&#x0E19;</span><span class="th-s" id="td2"></span></div>
<div class="tw"><table><thead><tr>
<th>&#x0E0A;&#x0E37;&#x0E48;&#x0E2D;</th><th>&#x0E04;&#x0E30;&#x0E41;&#x0E19;&#x0E19;&#x0E23;&#x0E27;&#x0E21;</th><th>Extra Like</th>
<th>Home Chef</th><th>Foodie Reviewer</th><th>Lunch Box Art</th>
<th>Extra &#x0E40;&#x0E1E;&#x0E37;&#x0E48;&#x0E2D;&#x0E19;&#x0E15;&#x0E32;&#x0E21;</th><th>Extra &#x0E42;&#x0E08;&#x0E17;&#x0E22;&#x0E4C;</th>
<th>&#x2764;&#xFE0F; Like</th><th>&#x0E42;&#x0E1E;&#x0E2A;&#x0E15;&#x0E4C;</th>
</tr></thead><tbody id="sb"></tbody></table></div></div>
<div class="sec"><div class="th"><span class="th-t">&#x1F35C; &#x0E23;&#x0E32;&#x0E22;&#x0E01;&#x0E32;&#x0E23;&#x0E40;&#x0E21;&#x0E19;&#x0E39;</span><span class="th-s" id="mc2"></span></div>
<div class="tw">
<div class="fbar" id="fb"><span class="fl">&#x0E01;&#x0E23;&#x0E2D;&#x0E07;&#x0E0A;&#x0E37;&#x0E48;&#x0E2D;:</span></div>
<table><thead><tr><th>&#x0E27;&#x0E31;&#x0E19;&#x0E17;&#x0E35;&#x0E48;</th><th>&#x0E0A;&#x0E37;&#x0E48;&#x0E2D;</th><th>&#x0E40;&#x0E21;&#x0E19;&#x0E39;</th><th>&#x0E2B;&#x0E21;&#x0E27;&#x0E14;</th><th>&#x2764;&#xFE0F;</th></tr></thead>
<tbody id="mb"></tbody></table>
</div></div>
<div class="mbg" id="mbg" onclick="if(event.target===this)this.classList.remove('open')">
<div class="modal"><button class="mc" onclick="document.getElementById('mbg').classList.remove('open')">&#x2715;</button>
<h2 id="mh"></h2><div id="md"></div></div></div>
<script>
""" + raw_js + "\n" + weeks_js + """
function iso(s){var m=s.trim().match(/^(\\d{1,2})\\/(\\d{1,2})\\/(\\d{4})$/);return m?m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0'):null;}
function dp(d){var p=d.split('-');return p[2]+'/'+p[1]+'/'+p[0];}
function sl(el){var v=el.value.replace(/\\D/g,'');if(v.length>2)v=v.slice(0,2)+'/'+v.slice(2);if(v.length>5)v=v.slice(0,5)+'/'+v.slice(5);el.value=v;}
function tg(r){var t='';if(r.homeChef)t+='<span class="tag hc">Home Chef</span>';if(r.foodie)t+='<span class="tag fd">Foodie</span>';if(r.lunchbox)t+='<span class="tag lb">Lunch Box</span>';if(r.extraTask)t+='<span class="tag et">Extra \\u0E42\\u0E08\\u0E17\\u0E22\\u0E4C</span>';if(r.extraFriend)t+='<span class="tag ef">Extra \\u0E40\\u0E1E\\u0E37\\u0E48\\u0E2D\\u0E19</span>';return t||'<span class="z">-</span>';}
function vl(x){return x?x:'<span class="z">-</span>';}
function calc(f0,t0){
  var f=new Date(f0),t=new Date(t0);
  var fil=RAW.filter(function(r){var d=new Date(r.date);return d>=f&&d<=t;});
  var map={};
  fil.forEach(function(r){var n=r.name;if(!map[n])map[n]={name:n,hc:0,fd:0,lb:0,ef:0,et:0,lk:0,posts:0,ent:[]};map[n].hc+=r.homeChef;map[n].fd+=r.foodie;map[n].lb+=r.lunchbox;map[n].ef+=r.extraFriend;map[n].et+=r.extraTask;map[n].lk+=r.countLike;map[n].posts++;map[n].ent.push(r);});
  var ps=Object.values(map);
  var sl2=[].concat(ps).sort(function(a,b){return b.lk-a.lk;});
  ps.forEach(function(p){var rk=sl2.findIndex(function(x){return x.name===p.name;});p.el=p.lk===0?0:rk===0?5:rk===1?3:1;p.total=p.hc+p.fd+p.lb+p.ef+p.et+p.el;});
  ps.sort(function(a,b){return b.total-a.total;});
  return {ps:ps,fil:fil};
}
function go(){
  var f0=iso(document.getElementById('f').value),t0=iso(document.getElementById('t').value);
  document.getElementById('mg').textContent=(!f0||!t0)?'\\u0E23\\u0E39\\u0E1B\\u0E41\\u0E1A\\u0E1A\\u0E44\\u0E21\\u0E48\\u0E16\\u0E39\\u0E01':'';
  if(!f0||!t0||f0>t0)return;
  document.getElementById('td2').textContent=document.getElementById('f').value+' - '+document.getElementById('t').value;
  var r=calc(f0,t0);
  window._pm={};r.ps.forEach(function(p){window._pm[p.name]=p;});
  dPod(r.ps);dCards(r.ps);dScore(r.ps);dMenu(r.fil);hl(f0,t0);
}
function dPod(ps){
  if(!ps.length){document.getElementById('pod').innerHTML='';return;}
  var top=ps.slice(0,3),ord=top.length>=3?[top[1],top[0],top[2]]:top,ri=top.length>=3?[1,0,2]:[0,1,2];
  var md=['\\uD83E\\uDD47','\\uD83E\\uDD48','\\uD83E\\uDD49'],rc=['p1','p2','p3'];
  document.getElementById('pod').innerHTML=ord.map(function(p,i){return '<div class="pc '+rc[ri[i]]+'"><div class="pm">'+md[ri[i]]+'</div><div class="pn">'+p.name+'</div><div class="pp">'+p.total+' pts</div></div>';}).join('');
}
function dCards(ps){
  var tp=ps.reduce(function(s,p){return s+p.posts;},0),tl=ps.reduce(function(s,p){return s+p.lk;},0),ml=ps.reduce(function(a,b){return a.lk>b.lk?a:b;},{name:'-',lk:0});
  document.getElementById('cds').innerHTML='<div class="card"><div class="cv">'+ps.length+'</div><div class="cl">\\u0E04;\\u0E19\\u0E40\\u0E02;\\u0E49\\u0E32\\u0E23;\\u0E48\\u0E27;\\u0E21</div></div>'+'<div class="card"><div class="cv">'+tp+'</div><div class="cl">\\u0E42\\u0E1E\\u0E2A\\u0E15\\u0E4C\\u0E23\\u0E27\\u0E21</div></div>'+'<div class="card"><div class="cv">'+tl+'</div><div class="cl">Like \\u0E23\\u0E27\\u0E21</div></div>'+'<div class="card"><div class="cv">'+ml.name+'</div><div class="cl">Like \\u0E2A\\u0E39\\u0E07\\u0E2A\\u0E38\\u0E14 ('+ml.lk+')</div></div>';
}
function dScore(ps){
  var rc=['r1','r2','r3'];
  if(!ps.length){document.getElementById('sb').innerHTML='<tr><td colspan="10" class="nd">-</td></tr>';return;}
  document.getElementById('sb').innerHTML=ps.map(function(p,i){return '<tr><td><span class="rb '+(rc[i]||'rn')+'">'+(i+1)+'</span><span class="ck" onclick="det(\\'' +p.name.replace(/\\\\/g,'\\\\\\\\').replace(/'/g,"\\\\'")+'\\')">' +p.name+'</span></td><td class="pts">'+p.total+'</td><td class="lk">'+vl(p.el)+'</td><td>'+vl(p.hc)+'</td><td>'+vl(p.fd)+'</td><td>'+vl(p.lb)+'</td><td>'+vl(p.ef)+'</td><td>'+vl(p.et)+'</td><td class="lk">'+p.lk+'</td><td>'+p.posts+'</td></tr>';}).join('');
}
var _mf=[];
function dMenu(fil){
  _mf=fil;
  var ns=['\\u0E17\\u0E31\\u0E49\\u0E07\\u0E2B\\u0E21\\u0E14'].concat([...new Set(fil.map(function(r){return r.name;}))].sort());
  document.getElementById('fb').innerHTML='<span class="fl">\\u0E01\\u0E23\\u0E2D\\u0E07\\u0E0A\\u0E37\\u0E48\\u0E2D:</span>'+ns.map(function(n,i){return '<button class="np'+(i===0?' all on':'')+'" onclick="fm(this,\\'' +n.replace(/'/g,"\\\\'")+'\\')">'+ n+'</button>';}).join('');
  am(null);
}
function fm(el,n){document.querySelectorAll('.np').forEach(function(p){p.classList.remove('on');});el.classList.add('on');am(n==='\\u0E17\\u0E31\\u0E49\\u0E07\\u0E2B\\u0E21\\u0E14'?null:n);}
function am(name){
  var d=name?_mf.filter(function(r){return r.name===name;}):_mf;
  d=[].concat(d).sort(function(a,b){return b.date.localeCompare(a.date);});
  document.getElementById('mc2').textContent=d.length+' \\u0E23\\u0E32\\u0E22\\u0E01\\u0E32\\u0E23';
  document.getElementById('mb').innerHTML=d.length?d.map(function(r){return '<tr><td class="td">'+dp(r.date)+'</td><td class="tn">'+r.name+'</td><td>'+(r.menu||'-')+'</td><td>'+tg(r)+'</td><td class="lk">'+r.countLike+'</td></tr>';}).join(''):'<tr><td colspan="5" class="nd">-</td></tr>';
}
function det(name){var p=(window._pm||{})[name];if(!p)return;document.getElementById('mh').textContent=p.name;document.getElementById('md').innerHTML=p.ent.map(function(r){return '<div class="dr"><span class="dd">'+dp(r.date)+'</span><span class="dm">'+(r.menu||'-')+'</span>'+tg(r)+'<span class="dl">&#x2764; '+r.countLike+'</span></div>';}).join('');document.getElementById('mbg').classList.add('open');}
function mkT(){document.getElementById('tabs').innerHTML=WEEKS.map(function(w,i){return '<button class="tab'+(i===0?' all':'')+'" id="tab'+i+'" onclick="sw('+i+')">'+ w.label+'</button>';}).join('');}
function sw(i){document.getElementById('f').value=dp(WEEKS[i].from);document.getElementById('t').value=dp(WEEKS[i].to);go();}
function hl(f0,t0){WEEKS.forEach(function(w,i){var e=document.getElementById('tab'+i);if(e)e.classList.toggle('on',w.from===f0&&w.to===t0);});}
mkT();sw(1);
</script></body></html>"""

out = os.path.join(BASE, 'scores.html')
with open(out,'w',encoding='utf-8') as f:
    f.write(html)
print('OK', len(html), 'chars written to', out)
