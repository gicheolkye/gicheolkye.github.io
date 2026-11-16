/**
 * 시험 시간표 학생 조회 — Google Apps Script (코드 파일 하나로 끝)
 *
 * 쓰는 법
 *  1) 시험 시간표 생성기에서 받은 '학생조회용_시간표.xlsx'를 구글 시트로 가져오기
 *     (파일 > 가져오기 > 업로드 > '스프레드시트 바꾸기')
 *  2) 그 시트에서 [확장 프로그램 > Apps Script]를 열고, 이 코드를 통째로 붙여넣고 저장
 *  3) [배포 > 새 배포 > 유형: 웹 앱] → 실행 사용자: 나 / 액세스 권한: 모든 사용자 → 배포
 *  4) 나온 웹 앱 주소를 학생들에게 안내
 *
 *  - 학생은 학번·이름(확인코드를 쓰면 확인코드까지)이 모두 맞아야 '자기 것만' 볼 수 있습니다.
 *  - 시트 자체는 공유하지 않으므로 학생이 다른 학생 정보를 열람할 수 없습니다.
 *  - '설정' 시트의 공개 값을 N으로 바꾸면 바로 조회가 막힙니다.
 *  - 다음 시험 때는 1)처럼 새 파일로 '스프레드시트 바꾸기'만 하면 같은 주소로 계속 쓸 수 있습니다.
 */

var MAX_FAIL = 5;      // 같은 학번으로 이만큼 틀리면
var LOCK_MIN = 10;     // 이 시간(분) 동안 조회 막기

function doGet() {
  var cfg = readConfig_();
  var html = PAGE_HTML
    .replace(/__TITLE__/g, esc_(cfg.title))
    .replace('__NOTICE__', esc_(cfg.notice).replace(/\n/g, '<br>'))
    .replace('__USECODE__', usesCode_() ? '1' : '0')
    .replace('__OPEN__', cfg.open ? '1' : '0');
  return HtmlService.createHtmlOutput(html)
    .setTitle(cfg.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function readConfig_() {
  var cfg = { title: '시험 시간표 조회', notice: '', open: true };
  var sh = SpreadsheetApp.getActive().getSheetByName('설정');
  if (!sh) return cfg;
  sh.getDataRange().getDisplayValues().forEach(function (r) {
    var k = String(r[0]).trim(), v = String(r[1] || '').trim();
    if (k === '제목' && v) cfg.title = v;
    if (k === '안내') cfg.notice = v;
    if (k === '공개') cfg.open = !/^(n|no|아니오|x|비공개)$/i.test(v);
  });
  return cfg;
}

function table_() {
  var sh = SpreadsheetApp.getActive().getSheetByName('시간표');
  if (!sh) throw new Error("'시간표' 시트를 찾을 수 없습니다.");
  var v = sh.getDataRange().getDisplayValues();
  var h = v[0].map(function (x) { return String(x).trim(); });
  return { values: v, h: h, iId: h.indexOf('학번'), iNm: h.indexOf('이름'), iCd: h.indexOf('확인코드') };
}

function usesCode_() {
  try {
    var t = table_();
    if (t.iCd < 0) return false;
    for (var r = 1; r < t.values.length; r++) if (String(t.values[r][t.iCd]).trim()) return true;
  } catch (e) {}
  return false;
}

/** 페이지에서 google.script.run.lookup(학번, 이름, 확인코드)로 부름 */
function lookup(id, name, code) {
  id = String(id || '').replace(/\D/g, '');
  name = String(name || '').replace(/\s/g, '');
  code = String(code || '').replace(/\s/g, '');
  if (!readConfig_().open) return { ok: false, msg: '지금은 시간표 조회 기간이 아닙니다.' };
  if (!/^\d{5}$/.test(id)) return { ok: false, msg: '학번 5자리를 입력해 주세요.' };
  if (!name) return { ok: false, msg: '이름을 입력해 주세요.' };

  var cache = CacheService.getScriptCache(), fk = 'fail_' + id;
  var fails = Number(cache.get(fk) || 0);
  if (fails >= MAX_FAIL) return { ok: false, msg: '입력이 여러 번 틀려서 ' + LOCK_MIN + '분 동안 조회할 수 없습니다. 잠시 뒤 다시 해 보세요.' };

  var t = table_(), row = null;
  for (var r = 1; r < t.values.length; r++) {
    if (String(t.values[r][t.iId]).replace(/\D/g, '') === id) { row = t.values[r]; break; }
  }
  var need = t.iCd >= 0 && row && String(row[t.iCd]).trim() !== '';
  var bad = !row || String(row[t.iNm]).replace(/\s/g, '') !== name || (need && String(row[t.iCd]).trim() !== code);
  if (bad) {
    cache.put(fk, String(fails + 1), LOCK_MIN * 60);
    Utilities.sleep(700);
    return { ok: false, msg: '입력한 정보와 일치하는 학생이 없습니다. 학번·이름' + (usesCode_() ? '·확인코드' : '') + '를 다시 확인해 주세요.' };
  }
  cache.remove(fk);
  var first = Math.max(t.iId, t.iNm, t.iCd) + 1, items = [];
  for (var c = first; c < t.h.length; c++) if (t.h[c]) items.push({ period: t.h[c], text: String(row[c] || '') });
  return { ok: true, id: id, name: String(row[t.iNm]), items: items };
}

function esc_(s) {
  return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
}

var PAGE_HTML = '<!DOCTYPE html>\n<html lang="ko"><head><meta charset="utf-8"><base target="_top">\n' +
'<style>\n' +
':root{--ac:#1f5c4a;--soft:#e6efe9;--ink:#1d2321;--mu:#68706c;--line:#e2e0d9;--hl:#fff36b}\n' +
'*{box-sizing:border-box}body{margin:0;background:#f6f5f1;color:var(--ink);font-family:"Noto Sans KR","Apple SD Gothic Neo","Malgun Gothic",system-ui,sans-serif;font-size:16px;line-height:1.5}\n' +
'header{background:var(--ac);color:#fff;padding:22px 18px 18px}header h1{margin:0;font-size:20px;letter-spacing:-.3px}header p{margin:6px 0 0;font-size:13.5px;opacity:.85}\n' +
'main{max-width:560px;margin:0 auto;padding:16px}\n' +
'.card{background:#fff;border:1px solid var(--line);border-radius:12px;padding:16px;margin-bottom:14px}\n' +
'label{display:block;font-size:13.5px;color:var(--mu);margin:10px 0 4px}\n' +
'input{width:100%;font:inherit;padding:11px 12px;border:1px solid #cfd3cf;border-radius:8px;font-size:17px}\n' +
'button{width:100%;margin-top:14px;font:inherit;font-weight:700;font-size:16px;padding:12px;border:none;border-radius:8px;background:var(--ac);color:#fff}\n' +
'button:disabled{opacity:.6}\n' +
'.msg{margin-top:10px;color:#a4461d;font-size:14px;min-height:1em}\n' +
'.who{display:flex;justify-content:space-between;align-items:baseline}.who b{font-size:19px}.who a{font-size:13.5px;color:var(--ac)}\n' +
'.day h2{font-size:15px;margin:0 0 8px;color:var(--ac)}\n' +
'.p{display:flex;gap:12px;padding:10px 0;border-top:1px solid var(--line)}.p:first-of-type{border-top:none}\n' +
'.p .no{flex:0 0 52px;font-weight:700;color:var(--mu);font-size:14px;padding-top:2px}\n' +
'.p .s{font-weight:700;font-size:17px}.p .r{font-size:14.5px}.p .seat{font-size:14.5px;color:var(--ac);font-weight:700}\n' +
'.p.miss .s{color:#a3a39c;font-weight:400;font-size:15px}\n' +
'.p.wait{background:var(--hl);margin:0 -16px;padding:10px 16px}.p.wait .s{color:var(--ink);font-weight:700}\n' +
'.closed{text-align:center;color:var(--mu);padding:30px 10px}\n' +
'.foot{font-size:12.5px;color:var(--mu);text-align:center;margin-top:8px}\n' +
'</style></head><body>\n' +
'<header><h1>__TITLE__</h1><p>__NOTICE__</p></header>\n' +
'<main>\n' +
'<div class="card" id="form">\n' +
'  <label for="id">학번 (5자리)</label><input id="id" inputmode="numeric" maxlength="5" autocomplete="off" placeholder="예: 20101">\n' +
'  <label for="nm">이름</label><input id="nm" autocomplete="off" placeholder="예: 홍길동">\n' +
'  <div id="cdBox" style="display:none"><label for="cd">확인코드 (담임 선생님께 받은 4자리)</label><input id="cd" inputmode="numeric" maxlength="8" autocomplete="off"></div>\n' +
'  <button id="go">내 시간표 보기</button>\n' +
'  <div class="msg" id="msg"></div>\n' +
'</div>\n' +
'<div id="out"></div>\n' +
'<p class="foot">입력한 정보는 시간표 확인에만 쓰이고 저장되지 않습니다.</p>\n' +
'</main>\n' +
'<script>\n' +
'var USECODE = "__USECODE__" === "1", OPEN = "__OPEN__" === "1";\n' +
'function $(s){return document.getElementById(s);}\n' +
'function esc(s){return String(s).replace(/[&<>"]/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;"}[c];});}\n' +
'if (USECODE) $("cdBox").style.display = "block";\n' +
'if (!OPEN) { $("form").innerHTML = "<div class=closed>지금은 시간표 조회 기간이 아닙니다.</div>"; }\n' +
'function go(){\n' +
'  var b=$("go"); b.disabled=true; b.textContent="찾는 중…"; $("msg").textContent="";\n' +
'  google.script.run.withSuccessHandler(function(res){\n' +
'    b.disabled=false; b.textContent="내 시간표 보기";\n' +
'    if(!res.ok){ $("msg").textContent=res.msg; return; }\n' +
'    show(res);\n' +
'  }).withFailureHandler(function(e){ b.disabled=false; b.textContent="내 시간표 보기"; $("msg").textContent="잠시 후 다시 시도해 주세요."; }).lookup($("id").value,$("nm").value,USECODE?$("cd").value:"");\n' +
'}\n' +
'function show(res){\n' +
'  var days=[], map={};\n' +
'  res.items.forEach(function(it){ var sp=it.period.indexOf(" "); var d=sp>0?it.period.slice(0,sp):it.period, p=sp>0?it.period.slice(sp+1):"";\n' +
'    if(!map[d]){map[d]=[];days.push(d);} map[d].push({p:p,t:it.text}); });\n' +
'  var h="<div class=card><div class=who><b>"+esc(res.id)+" "+esc(res.name)+"</b><a href=\\"#\\" onclick=\\"reset();return false;\\">다른 학생 조회</a></div></div>";\n' +
'  days.forEach(function(d){\n' +
'    h+="<div class=\\"card day\\"><h2>"+esc(d)+"</h2>";\n' +
'    map[d].forEach(function(x){\n' +
'      var L=x.t.split("\\n"), cls="p", body="";\n' +
'      if(L[0].indexOf("미응시")>=0){\n' +
'        if(L[1]){ cls+=" wait"; body="<div class=s>응시 과목 없음 · 학교에서 대기</div><div class=r>"+esc(L[1])+"</div>"; }\n' +
'        else { cls+=" miss"; body="<div class=s>응시 과목 없음</div>"; }\n' +
'      } else if(L[0]) {\n' +
'        body="<div class=s>"+esc(L[0])+"</div>"+(L[1]?"<div class=r>"+esc(L[1])+"</div>":"")+(L[2]?"<div class=seat>"+esc(L[2])+"</div>":"");\n' +
'      } else { cls+=" miss"; body="<div class=s>-</div>"; }\n' +
'      h+="<div class=\\""+cls+"\\"><div class=no>"+esc(x.p)+"</div><div>"+body+"</div></div>";\n' +
'    });\n' +
'    h+="</div>";\n' +
'  });\n' +
'  $("form").style.display="none"; $("out").innerHTML=h; window.scrollTo(0,0);\n' +
'}\n' +
'function reset(){ $("out").innerHTML=""; $("form").style.display="block"; $("id").value=""; $("nm").value=""; if($("cd"))$("cd").value=""; $("id").focus(); }\n' +
'$("go").onclick=go;\n' +
'["id","nm","cd"].forEach(function(k){ var e=$(k); if(e) e.addEventListener("keydown",function(ev){ if(ev.key==="Enter") go(); }); });\n' +
'</script></body></html>';
