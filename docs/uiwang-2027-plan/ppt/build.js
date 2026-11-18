const pptxgen = require('pptxgenjs');
const { applyTheme } = require('/root/.claude/skills/synced/207274f0-2193-4748-98af-cca2d4b37692_ca0453cc-6b98-44dc-88c0-e61e491c982b/pptx/scripts/apply_theme.js');
const THEME = { name: 'Uiwang Care', headFontFace: '맑은 고딕', bodyFontFace: '맑은 고딕',
  colors: { dk1:'1D2B3A', lt1:'FFFFFF', dk2:'1B3A5C', lt2:'EEF3F8', accent1:'1B3A5C', accent2:'E58A1F', accent3:'2F5D8A', accent4:'9FB6CC', accent5:'DCE7F2', accent6:'5B6F84', hlink:'2F5D8A', folHlink:'5B6F84' } };
const HX = { navy:'1B3A5C', orange:'E58A1F', mid:'2F5D8A', line:'9FB6CC', pale:'DCE7F2', grey:'5B6F84', text:'1D2B3A' };
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = '2026년 제2차 의왕시 통합지원협의체';
const C = pres.SchemeColor;
const FOOT = '2026년 제2차 의왕시 통합지원협의체';

pres.defineSlideMaster({ title:'COVER', background:{ color:C.text2 },
  objects:[ { placeholder:{ options:{ name:'title', type:'title', x:0.8, y:2.2, w:11.7, h:1.6, fontSize:40, bold:true, color:C.background1, valign:'top' }, text:'' } },
            { placeholder:{ options:{ name:'body', type:'body', x:0.8, y:4.0, w:11.7, h:1.6, fontSize:20, color:C.accent5, valign:'top' }, text:'' } } ] });
pres.defineSlideMaster({ title:'SECTION', background:{ color:C.text2 },
  objects:[ { placeholder:{ options:{ name:'title', type:'title', x:0.8, y:3.0, w:11.7, h:1.2, fontSize:36, bold:true, color:C.background1, valign:'top' }, text:'' } },
            { placeholder:{ options:{ name:'body', type:'body', x:0.8, y:2.2, w:11.7, h:0.7, fontSize:20, bold:true, color:C.accent2, valign:'bottom' }, text:'' } } ] });
pres.defineSlideMaster({ title:'CONTENT', background:{ color:C.background1 },
  margin:[0.5,0.6,0.6,0.6],
  objects:[ { placeholder:{ options:{ name:'title', type:'title', x:0.6, y:0.35, w:12.1, h:0.8, fontSize:28, bold:true, color:C.text2, valign:'middle' }, text:'' } },
            { text:{ text:FOOT, options:{ x:0.6, y:7.0, w:8, h:0.3, fontSize:10, color:C.accent6 } } } ],
  slideNumber:{ x:12.2, y:7.0, w:0.5, h:0.3, fontSize:10, color:C.accent6, align:'right' } });

let sec='';
function S(master, title){ const s=pres.addSlide({ masterName:master, sectionTitle:sec }); s.addText(title,{ placeholder:'title' }); return s; }
function txt(s,t,o){ s.addText(t,Object.assign({ isTextBox:true, fontSize:14, color:C.text1, margin:0.05, valign:'top' },o)); }
function box(s,x,y,w,h,fill,name){ s.addShape(pres.shapes.ROUNDED_RECTANGLE,{ x,y,w,h, rectRadius:0.12, fill:{ color:fill }, line:{ color:fill }, objectName:name }); }
function stat(s,x,y,w,h,num,label,sub,dark){
  box(s,x,y,w,h, dark?C.text2:C.background2, 'stat-'+label);
  txt(s,num,{ x:x+0.25,y:y+0.25,w:w-0.5,h:1.0, fontSize:40, bold:true, color: dark?C.background1:C.text2 });
  txt(s,label,{ x:x+0.25,y:y+1.3,w:w-0.5,h:0.45, fontSize:16, bold:true, color: dark?C.accent5:C.text1 });
  if(sub) txt(s,sub,{ x:x+0.25,y:y+1.75,w:w-0.5,h:h-1.85, fontSize:13, color: dark?C.accent5:C.accent6 });
}
function bullets(s,items,o){ s.addText(items.map((t,i)=>({ text:t, options:{ bullet:true, breakLine:i<items.length-1 } })), Object.assign({ isTextBox:true, fontSize:15, color:C.text1, paraSpaceAfter:8, valign:'top', margin:0.05 },o)); }
function head(text){ return { text, options:{ bold:true, color:C.background1, fill:{ color:C.text2 }, align:'center', valign:'middle' } }; }
function table(s,rows,o){ s.addTable(rows,Object.assign({ fontSize:14, color:C.text1, border:{ type:'solid', pt:0.75, color:HX.line }, valign:'middle', margin:[0.05,0.08,0.05,0.08] },o)); }
function num(s,x,y,n,dark){ s.addShape(pres.shapes.OVAL,{ x,y,w:0.5,h:0.5, fill:{ color: dark?C.accent2:C.text2 }, line:{ color: dark?C.accent2:C.text2 } }); txt(s,String(n),{ x,y,w:0.5,h:0.5, fontSize:16, bold:true, color:C.background1, align:'center', valign:'middle', margin:0 }); }

// 1 cover
sec='표지'; pres.addSection({ title:sec });
let s=pres.addSlide({ masterName:'COVER', sectionTitle:sec });
s.addText('2027년도 통합지원 지역계획(안)',{ placeholder:'title' });
s.addText([{ text:'2026년 제2차 의왕시 통합지원협의체', options:{ breakLine:true } },{ text:'2026. 11. 26.  |  의왕시 복지정책과 통합돌봄팀' }],{ placeholder:'body' });
s.addNotes('2026년 제2차 의왕시 통합지원협의체 회의를 시작하겠습니다. 오늘은 2026년 추진현황을 보고드리고, 2027년도 통합지원 지역계획(안)을 심의해 주시기 바랍니다.');

// 2 agenda
sec='안건'; pres.addSection({ title:sec });
s=S('CONTENT','오늘의 안건');
[[ '보고안건 1','2026년 통합돌봄 추진현황','법 시행 첫해 성과와 1차 회의 후속 조치' ],[ '심의안건 1','2027년도 통합지원 지역계획(안)','지역 진단, 비전·목표, 핵심사업 12개, 성과지표, 예산' ]].forEach((a,i)=>{
  const x=0.6+i*6.2; box(s,x,1.6,5.9,4.6, i?C.text2:C.background2,'agenda'+i);
  txt(s,a[0],{ x:x+0.4,y:2.0,w:5.1,h:0.5, fontSize:18, bold:true, color:C.accent2 });
  txt(s,a[1],{ x:x+0.4,y:2.6,w:5.1,h:1.4, fontSize:28, bold:true, color: i?C.background1:C.text2 });
  txt(s,a[2],{ x:x+0.4,y:4.3,w:5.1,h:1.4, fontSize:16, color: i?C.accent5:C.accent6 });
});
s.addNotes('안건은 보고안건 1건, 심의안건 1건입니다.');

// 3 section
sec='보고안건 1'; pres.addSection({ title:sec });
s=pres.addSlide({ masterName:'SECTION', sectionTitle:sec }); s.addText('2026년 통합돌봄 추진현황',{ placeholder:'title' }); s.addText('보고안건 1',{ placeholder:'body' });

// 4 key stats
s=S('CONTENT','법 시행 첫해, 276명에게 서비스 740건을 연계');
stat(s,0.6,1.5,2.85,2.55,'390명','신청','’26. 10. 31. 기준');
stat(s,3.68,1.5,2.85,2.55,'276명','통합지원 대상자','1차 회의 목표 46명의 6배',true);
stat(s,6.76,1.5,2.85,2.55,'740건','서비스 연계','대상자 1인당 2.7건');
stat(s,9.84,1.5,2.85,2.55,'340명','의왕살핌 이용','일상 138 · 건강 103 · 내집 99');
txt(s,'대상자 유형',{ x:0.6,y:4.4,w:4,h:0.4, fontSize:16, bold:true, color:C.text2 });
s.addChart(pres.charts.BAR,[{ name:'대상자', labels:['재가급여자','기타','등급외자(기각 포함)','퇴원환자','고령장애인','맞돌 중점군'], values:[103,55,52,43,22,1] }],
  { x:0.6,y:4.8,w:6.6,h:2.1, barDir:'bar', chartColors:[HX.navy], showValue:true, dataLabelPosition:'outEnd', dataLabelFontSize:11, dataLabelColor:HX.text, catAxisLabelColor:HX.grey, catAxisLabelFontSize:11, valAxisHidden:true, valGridLine:{ style:'none' }, catGridLine:{ style:'none' }, showLegend:false, catAxisOrientation:'maxMin', catAxisLabelFontFace:'+mn-lt', dataLabelFontFace:'+mn-lt' });
txt(s,'연계 분야',{ x:7.6,y:4.4,w:4,h:0.4, fontSize:16, bold:true, color:C.text2 });
s.addChart(pres.charts.BAR,[{ name:'연계', labels:['일상돌봄','건강관리·예방','주거','보건의료','장기요양','기타'], values:[241,143,131,127,30,68] }],
  { x:7.6,y:4.8,w:5.1,h:2.1, barDir:'bar', chartColors:[HX.orange], showValue:true, dataLabelPosition:'outEnd', dataLabelFontSize:11, dataLabelColor:HX.text, catAxisLabelColor:HX.grey, catAxisLabelFontSize:11, valAxisHidden:true, valGridLine:{ style:'none' }, catGridLine:{ style:'none' }, showLegend:false, catAxisOrientation:'maxMin', catAxisLabelFontFace:'+mn-lt', dataLabelFontFace:'+mn-lt' });
s.addNotes('법 시행 첫해 신청 390명 중 276명을 통합지원 대상자로 결정했습니다. 3월 1차 회의에서 말씀드린 목표 46명의 6배입니다. 서비스는 740건, 1인당 2.7건을 연계했고, 일상돌봄과 건강관리 연계가 가장 많았습니다.');

// 5 discovery
s=S('CONTENT','집중홍보와 동별 목표관리로 10월 신청이 5배 늘었습니다');
s.addChart(pres.charts.BAR,[{ name:'신청', labels:['1~9월 월평균','10월'], values:[27,147] }],
  { x:0.6,y:1.5,w:4.6,h:4.6, barDir:'col', chartColors:[HX.line,HX.orange], showValue:true, dataLabelPosition:'outEnd', dataLabelFontSize:16, dataLabelColor:HX.text, catAxisLabelColor:HX.grey, catAxisLabelFontSize:14, valAxisHidden:true, valGridLine:{ style:'none' }, showLegend:false, showTitle:true, title:'월 신청(명)', titleFontSize:14, titleColor:HX.navy, catAxisLabelFontFace:'+mn-lt', dataLabelFontFace:'+mn-lt', titleFontFace:'+mn-lt' });
s.addChart(pres.charts.BAR,[{ name:'9월', labels:['내손2동','청계동','오전동','시 전체'], values:[57.9,64.4,64.5,74.8] },{ name:'10월', labels:['내손2동','청계동','오전동','시 전체'], values:[207.2,91.7,88.3,119.4] }],
  { x:5.5,y:1.5,w:7.2,h:4.6, barDir:'col', barGrouping:'clustered', chartColors:[HX.line,HX.navy], showValue:true, dataLabelPosition:'outEnd', dataLabelFontSize:12, dataLabelColor:HX.text, catAxisLabelColor:HX.grey, catAxisLabelFontSize:14, valAxisHidden:true, valGridLine:{ style:'none' }, showLegend:true, legendPos:'t', legendFontSize:12, showTitle:true, title:'노인 1만 명당 신청(명), 9월 → 10월', titleFontSize:14, titleColor:HX.navy, catAxisLabelFontFace:'+mn-lt', dataLabelFontFace:'+mn-lt', titleFontFace:'+mn-lt', legendFontFace:'+mn-lt' });
bullets(s,['2차 집중홍보: 동 지역사회보장협의체 160명, 장기요양기관 5개소','동별 발굴목표 시달, 실적이 낮은 동은 시가 직접 발굴·접수 지원','건보공단 빅데이터 명단 454명 동별 확인·조사'],{ x:0.6,y:6.2,w:12.1,h:0.8, fontSize:14, paraSpaceAfter:2 });
s.addNotes('운영 초기에는 본인 신청 위주였습니다. 9~10월 2차 집중홍보와 동별 목표관리를 함께 추진해 10월 한 달 신청이 147명으로, 1~9월 월평균 27명의 5배가 되었습니다. 실적이 낮던 내손2동, 청계동, 오전동이 크게 개선되었습니다.');

// 6 services
s=S('CONTENT','지역특화서비스 「의왕살핌」으로 돌봄 공백을 메웠습니다');
table(s,[[head('구분'),head('서비스'),head('이용(명)'),head('제공기관')],
 ['일상살핌','생활돌봄 · 동행돌봄 · 식사지원',{text:'138',options:{align:'center',bold:true}},'(사)한국커뮤니티연구원, 의왕시니어클럽'],
 ['건강살핌','방문의료 · 가정간호 · 방문재활',{text:'103',options:{align:'center',bold:true}},'방문진료 의원 14개소, 서울홈케어내과, 물리치료사협회'],
 ['내집살핌','주거환경개선',{text:'99',options:{align:'center',bold:true}},'미소나눔'],
 ['특화공모(8월)','건강돌봄가 · 맞춤형 영양식 · 동고동락',{text:'127',options:{align:'center',bold:true}},'한국커뮤니티연구원, 의왕시니어클럽, 아름채노인복지관']],
 { x:0.6,y:1.5,w:7.6,colW:[1.4,2.5,0.9,2.8], rowH:0.62, fontSize:13 });
box(s,8.6,1.5,4.1,2.55,C.text2,'homecare');
txt(s,'14개소',{ x:8.85,y:1.7,w:3.6,h:0.9, fontSize:40, bold:true, color:C.background1 });
txt(s,'방문진료 의원이 실제 진료 시작',{ x:8.85,y:2.65,w:3.6,h:0.45, fontSize:16, bold:true, color:C.accent5 });
txt(s,'의과 2 · 한의 12, 등록만 되어 있던 의원 참여',{ x:8.85,y:3.15,w:3.6,h:0.7, fontSize:13, color:C.accent5 });
box(s,8.6,4.3,4.1,2.45,C.background2,'hhc');
txt(s,'2개소',{ x:8.85,y:4.5,w:3.6,h:0.9, fontSize:40, bold:true, color:C.text2 });
txt(s,'재택의료센터 확보',{ x:8.85,y:5.45,w:3.6,h:0.45, fontSize:16, bold:true, color:C.text1 });
txt(s,'관내 한국한센복지협회 부설의원, 관외 서울홈케어내과의원',{ x:8.85,y:5.95,w:3.6,h:0.7, fontSize:13, color:C.accent6 });
bullets(s,['국가돌봄서비스 우선 연계 80건, 공백만 지역특화서비스로 보완','소득과 무관하게 신청, 중위소득 기준 자부담 차등'],{ x:0.6,y:5.85,w:7.6,h:0.95, fontSize:14, paraSpaceAfter:4 });
s.addNotes('공적 서비스가 닿지 않는 공백을 메우기 위해 의왕살핌 3대 패키지를 4월에 신설했고, 340명이 이용했습니다. 상반기에 확인한 치매·영양·건강 모니터링 공백은 8월 특화공모 3종으로 대응했습니다. 실태조사 당시 관내 공급이 없던 방문진료는 의원 14개소가 실제 진료를 시작했습니다.');

// 7 delivery
s=S('CONTENT','조례·조직·회의체를 갖추고 협업 체계를 만들었습니다');
const dl=[['조례','「의왕시 의료·요양 등 지역돌봄의 통합지원에 관한 조례」 시행(3. 27.)'],['조직·인력','통합돌봄팀 신설(2월, 복지·간호), 전담인력 배치율 100%, 동 전담인력 6명 배치(11. 30.)'],['통합지원협의체','의료·복지·공단·학계·시의회 18명, 보건의료 위원 4명'],['통합지원회의','15회 개최, 300건 심의(신규 231·변경 29·종결 40), 보건소 전 회의 참석'],['기관 협업','협약 제공기관 18개소, 퇴원환자 협약병원 3개소, 간담회(8월, 20개소 32명)']];
dl.forEach((d,i)=>{ const y=1.5+i*1.05; num(s,0.6,y+0.15,i+1); txt(s,d[0],{ x:1.3,y:y+0.1,w:2.6,h:0.6, fontSize:18, bold:true, color:C.text2, valign:'middle' }); txt(s,d[1],{ x:4.0,y:y+0.1,w:8.7,h:0.6, fontSize:15, valign:'middle' }); if(i<4) s.addShape(pres.shapes.LINE,{ x:1.3,y:y+0.92,w:11.4,h:0, line:{ color:HX.pale, width:1 } }); });
s.addNotes('법 시행 전에 조례를 마련했고, 복지와 간호 직렬을 함께 배치한 통합돌봄팀을 신설했습니다. 통합지원회의는 15회, 300건을 심의했고 보건소가 모든 회의에 참석했습니다.');

// 8 follow-up
s=S('CONTENT','1차 회의(3. 11.) 위원 의견을 이렇게 반영했습니다');
table(s,[[head('위원'),head('1차 회의 의견'),head('반영 결과')],
 ['이민영\n(의사회)','의료 사각지대에 방치된 분을 더 찾을 수 있는 길 마련','빅데이터 명단 454명 조사, 의뢰 경로 확대, 방문진료 14개소 진료 시작 → ’27년 85세 이상 1인세대 조사'],
 ['김범준\n(한림대성심병원)','의료계 협업 발굴, 의뢰 후 판정까지 시간 단축','퇴원환자 협약병원 3개소, 신규 대면·변경 서면 회의, 한 번 방문 조사 → ’27년 퇴원환자 의뢰 경로 이원화'],
 ['이정근\n(약사회)','예산 3억 사용처, 다제약물 사업 기재','의왕살핌 3대 패키지·특화공모 3종 운영, 수행기관을 의왕시약사회로 기재']],
 { x:0.6,y:1.5,w:12.1,colW:[2.0,3.8,6.3], rowH:[0.5,1.15,1.15,1.15] });
box(s,0.6,5.75,12.1,1.0,C.background2,'target');
txt(s,[{ text:'1차 회의 목표 ', options:{ color:C.text1 } },{ text:'46명', options:{ bold:true, color:C.text2 } },{ text:'  →  2026년 통합지원 대상자 ', options:{ color:C.text1 } },{ text:'276명', options:{ bold:true, color:C.accent2 } }],{ x:0.9,y:5.85,w:11.5,h:0.8, fontSize:22, valign:'middle' });
s.addNotes('1차 회의에서 주신 의견을 어떻게 반영했는지 보고드립니다. 이민영 위원님의 의료 사각지대 발굴, 김범준 위원님의 의료계 협업과 판정 기간 단축 의견은 2027년 계획에도 이어서 반영했습니다.');

// 9 gaps
s=S('CONTENT','남은 과제는 2027년 계획으로 이어집니다');
[['발굴','본인 신청 위주로 접수, 발굴경로별 실적 미집계','85세 이상 1인세대 조사, 빅데이터 반기 배분, 발굴경로 기록'],['퇴원환자','협약병원 공식 의뢰 0건(상급병원은 퇴원일 전날 확정)','병원 유형별 의뢰 경로 이원화, 연계수당'],['사후관리','의뢰가 늘면서 서비스 개시 확인 지연','동 전담인력 모니터링, 제공기록 월 1회 행복e음 입력']].forEach((g,i)=>{
  const y=1.55+i*1.75;
  box(s,0.6,y,1.8,1.45,C.text2,'g'+i); txt(s,g[0],{ x:0.6,y,w:1.8,h:1.45, fontSize:20, bold:true, color:C.background1, align:'center', valign:'middle' });
  box(s,2.6,y,4.6,1.45,C.background2,'gp'+i); txt(s,g[1],{ x:2.85,y,w:4.1,h:1.45, fontSize:15, valign:'middle' });
  s.addShape(pres.shapes.RIGHT_TRIANGLE,{ x:7.35,y:y+0.5,w:0.4,h:0.45, fill:{ color:C.accent2 }, line:{ color:C.accent2 }, rotate:90 });
  box(s,7.9,y,4.8,1.45,C.background1,'gn'+i); s.addShape(pres.shapes.ROUNDED_RECTANGLE,{ x:7.9,y,w:4.8,h:1.45, rectRadius:0.12, fill:{ color:C.background1 }, line:{ color:C.accent2, width:1.5 } });
  txt(s,g[2],{ x:8.15,y,w:4.3,h:1.45, fontSize:15, bold:true, color:C.text2, valign:'middle' });
});
txt(s,'2026년 한계',{ x:2.6,y:1.1,w:4,h:0.4, fontSize:13, color:C.accent6 }); txt(s,'2027년 대응',{ x:7.9,y:1.1,w:4,h:0.4, fontSize:13, color:C.accent2, bold:true });
s.addNotes('한계도 있었습니다. 발굴은 본인 신청 위주였고, 협약병원의 공식 의뢰는 한 건도 없었으며, 의뢰가 늘면서 서비스 시작 확인이 늦어졌습니다. 이 세 가지가 2027년 계획의 출발점입니다.');

// 10 section
sec='심의안건 1'; pres.addSection({ title:sec });
s=pres.addSlide({ masterName:'SECTION', sectionTitle:sec }); s.addText('2027년도 통합지원 지역계획(안)',{ placeholder:'title' }); s.addText('심의안건 1',{ placeholder:'body' });

// 11 diagnosis
s=S('CONTENT','노인은 젊은 편이지만, 돌봄이 필요한 층은 초고령입니다');
stat(s,0.6,1.5,3.9,2.4,'33,930명','노인인구(고령화율 20.3%)','청계·오전·부곡동에 64.2% 거주');
stat(s,4.7,1.5,3.9,2.4,'48.6%','핵심군 중 85세 이상','경기 42.8%, 31개 시군 중 3위',true);
stat(s,8.8,1.5,3.9,2.4,'1,046세대','85세 이상 1인세대','부곡 292 · 청계 211 · 내손1 177');
table(s,[[head('진단'),head('근거'),head('2027년 방향')],
 ['방문의료 이용 저조','수요 668명 대비 방문의료·가정간호 70명','발굴 대상자 건강살핌 집중 연계'],
 ['퇴원 후 지역 복귀 지원','사회적 입원 224명, 재입원 고위험 90명','의뢰 경로 이원화, 회복살핌'],
 ['장애인 건강 공백','장애인 건강주치의 무공급, 재활 12.1%(전국 33.7%)','참여 의원 발굴'],
 ['서비스 인지도','인지 52.9% → 이용 6.4%','찾아가는 설명회, 생활 밀착 홍보']],
 { x:0.6,y:4.15,w:12.1,colW:[2.8,5.3,4.0], rowH:0.5, fontSize:14 });
s.addNotes('우리 시 노인은 전체적으로 젊은 편이지만, 돌봄이 필요한 핵심군은 85세 이상이 절반 가까이로 경기도에서 세 번째로 높습니다. 그래서 기존 저소득·독거 중심 발굴로는 잡히지 않는 초고령 가구를 먼저 찾는 것을 2027년 첫 과제로 삼았습니다.');

// 12 vision
s=S('CONTENT','비전·목표와 4대 과업, 12개 핵심사업');
s.addImage({ path:'vision.png', x:0.95,y:1.3,w:11.4,h:5.72, altText:'2027년 비전 체계도: 비전, 목표 3개, 추진과업 4개와 핵심사업 12개' });
s.addNotes('비전은 살던 곳에서 건강하게, 함께 살피는 의왕입니다. 2026년 실행계획의 비전을 이어받아 발굴과 사후관리까지 함께하는 체계로 넓혔습니다. 목표 3개 아래 추진과업 4개, 핵심사업 12개로 구성했습니다.');

// 13 targets & discovery
s=S('CONTENT','2027년 목표 대상자 541명, 초고령 위험군부터 찾습니다');
box(s,0.6,1.5,3.5,1.95,C.text2,'t541'); txt(s,'541명',{ x:0.85,y:1.6,w:3,h:0.85, fontSize:40, bold:true, color:C.background1 }); txt(s,'2027년 목표 대상자',{ x:0.85,y:2.45,w:3,h:0.4, fontSize:16, bold:true, color:C.accent5 }); txt(s,'우선발굴필요군 421 + 기타 120',{ x:0.85,y:2.85,w:3,h:0.4, fontSize:13, color:C.accent5 });
table(s,[[head('우선발굴필요군 유형'),head('규모'),head('’27 목표')],
 ['장기요양 재가급여자','2,151','145'],['고령장애인','2,871','107'],['의료기관 퇴원환자','840','60'],['등급외자·맞돌 중점군','235','50'],['장기요양 기각·각하자','361','40'],['65세 미만 중증장애인','298','17'],['장기요양시설 퇴소자','12','2'],
 [{text:'계(중복 제거)',options:{bold:true}},{text:'6,506',options:{bold:true}},{text:'421',options:{bold:true,color:C.accent2}}]],
 { x:0.6,y:3.6,w:3.5,colW:[1.9,0.8,0.8], rowH:0.3, fontSize:11, margin:[0.02,0.06,0.02,0.06] });
txt(s,'초고령 위험군 선제발굴',{ x:4.5,y:1.5,w:8.2,h:0.45, fontSize:20, bold:true, color:C.text2 });
bullets(s,['기초연금 수급 85세 이상 1인세대 명단을 시가 상·하반기에 추출 → 동이 연중 방문조사','건보공단 빅데이터(입원위험 재가급여자, 등급외·기각자) 반기 1회 동별 배분','신청 시 발굴경로 기록, 동별 발굴률 월별 관리 → 발굴률 20% 이상'],{ x:4.5,y:2.0,w:8.2,h:1.9, fontSize:15 });
txt(s,'퇴원환자 의뢰 경로 이원화',{ x:4.5,y:3.95,w:8.2,h:0.45, fontSize:20, bold:true, color:C.text2 });
s.addImage({ path:'discharge.png', x:4.5,y:4.45,w:8.2,h:2.04, altText:'퇴원환자 의뢰 경로 이원화: 상급종합병원은 안내 후 동 신청, 종합병원은 시에 직접 의뢰' });
s.addNotes('2027년 목표 대상자는 541명입니다. 우선발굴필요군 6,506명 가운데 421명, 약 6%를 목표로 했습니다. 1차 회의 때 2,319명은 건보 지역진단 기준이고, 이번에는 건보공단 우선관리대상자 산출 자료 기준이라 규모가 다릅니다. 퇴원환자는 상급종합병원은 안내 후 동 신청, 종합병원은 시에 직접 의뢰하도록 경로를 나눴습니다.');

// 14 monitoring
s=S('CONTENT','동 전담인력이 발굴부터 모니터링까지 끝까지 살핍니다');
const steps=['발굴·신청','사전·자체조사','지원계획 수립','통합지원회의','서비스 연계','모니터링'];
steps.forEach((t,i)=>{ const x=0.6+i*2.05; box(s,x,1.6,1.85,1.1, i==5?C.accent2:C.text2,'st'+i); txt(s,t,{ x,y:1.6,w:1.85,h:1.1, fontSize:15, bold:true, color:C.background1, align:'center', valign:'middle' }); });
txt(s,'동 전담인력 1명이 전 과정을 일관 관리  ·  방문조사는 간호직과 2인 1조',{ x:0.6,y:2.85,w:12.1,h:0.45, fontSize:15, color:C.accent6, align:'center' });
table(s,[[head('구분'),head('대상'),head('주기'),head('방법')],
 [{text:'최초',options:{bold:true,align:'center'}},'신규 서비스 연계 대상자','지원계획 안내일로부터 1개월 이내','대면·유선'],
 [{text:'정기',options:{bold:true,align:'center'}},'일반 대상자','3개월마다 1회 이상','제공기록 확인 후 유선, 필요 시 대면']],
 { x:0.6,y:3.55,w:7.4,colW:[0.9,2.2,2.3,2.0], rowH:0.6, fontSize:13 });
box(s,8.4,3.55,4.3,2.9,C.background2,'mon');
txt(s,'80%',{ x:8.65,y:3.75,w:3.8,h:0.9, fontSize:40, bold:true, color:C.text2 });
txt(s,'모니터링 수행률 목표',{ x:8.65,y:4.65,w:3.8,h:0.45, fontSize:16, bold:true });
txt(s,'상태 변화 시 전담인력이 변경(안) 작성 → 통합지원회의 재상정',{ x:8.65,y:5.15,w:3.8,h:1.1, fontSize:13, color:C.accent6 });
bullets(s,['제공기관은 제공주기·최초 제공일·상태 변화를 매월 시에 제출, 시가 행복e음 입력','동 전담인력은 시스템에서 진행상황을 직접 확인'],{ x:0.6,y:5.85,w:7.4,h:0.95, fontSize:14, paraSpaceAfter:4 });
s.addNotes('2026년에 확인한 서비스 시작 확인 지연 문제를 해결하기 위해, 11월 30일 배치된 동 전담인력이 발굴부터 모니터링까지 한 사람이 끝까지 관리합니다. 방문조사는 간호직과 2인 1조로 나가 의료 필요를 함께 확인합니다.');

// 15 service expansion
s=S('CONTENT','의왕살핌 30% 확대, 돌봄 공백 서비스 3종 신규 도입');
table(s,[[head('서비스'),head('’26 실적'),head('’27 목표')],
 ['일상살핌','138명',{text:'180명',options:{bold:true,color:C.text2}}],['건강살핌','103명',{text:'134명',options:{bold:true,color:C.text2}}],['내집살핌','99명',{text:'129명',options:{bold:true,color:C.text2}}],['맞춤형 영양식(개편)','55명',{text:'72명',options:{bold:true,color:C.text2}}],['건강돌봄가','50명',{text:'65명',options:{bold:true,color:C.text2}}],['동고동락','22명',{text:'29명',options:{bold:true,color:C.text2}}],['퇴원환자 연계','의뢰 0건',{text:'10명',options:{bold:true,color:C.text2}}]],
 { x:0.6,y:1.5,w:5.6,colW:[2.4,1.6,1.6], rowH:0.55, fontSize:14 });
txt(s,'식사지원은 일반식을 기존 서비스로 연계하고, 퇴원·급성기 맞춤형 영양식으로 개편',{ x:0.6,y:6.0,w:5.6,h:0.8, fontSize:13, color:C.accent6 });
[['이동살핌','승강기 없는 주택의 거동불편자 병원 진료 이동 지원','30명'],['회복살핌 꾸러미','퇴원 초기·치매 대상 건강관리용품, 노인맞춤돌봄 단기집중과 함께','100명'],['AI 돌봄로봇','독거·안전확인 필요 대상 안부 확인','100명']].forEach((n,i)=>{
  const y=1.5+i*1.75; box(s,6.6,y,6.1,1.55,C.background2,'new'+i);
  txt(s,'신규',{ x:6.85,y:y+0.2,w:0.8,h:0.38, fontSize:12, bold:true, color:C.background1, fill:{ color:C.accent2 }, align:'center', valign:'middle', margin:0 });
  txt(s,n[0],{ x:7.8,y:y+0.12,w:3.2,h:0.55, fontSize:20, bold:true, color:C.text2, valign:'middle' });
  txt(s,n[2],{ x:11.0,y:y+0.12,w:1.5,h:0.55, fontSize:20, bold:true, color:C.accent2, align:'right', valign:'middle' });
  txt(s,n[1],{ x:6.85,y:y+0.75,w:5.6,h:0.7, fontSize:14 });
});
s.addNotes('일상·건강·내집살핌 이용자를 전년보다 30% 늘리고, 식사지원은 퇴원·급성기 맞춤형 영양식으로 바꿉니다. 새로 이동살핌, 회복살핌 꾸러미, AI 돌봄로봇 세 가지를 도입합니다. 장애인 건강주치의 참여 의원도 1개소 이상 발굴하겠습니다.');

// 16 ops & edu
s=S('CONTENT','운영기반·민관협력과 교육·홍보');
const ops=[['서비스 의뢰·회신 표준체계 구축','신규','표준서식 개정, 접수 확인 회신, 제공기록 월 1회 입력'],['전문기관 협업 정례화','확대','건보 빅데이터·한 번 방문 조사, 관외 협력기관(서울홈케어내과) 연계'],['통합지원회의 운영 내실화','개선','월 2회 이상, 신규 대면·변경 종결 서면, 보건소 참석']];
const edu=[['「의왕돌봄마당」 민관 역량강화','확대','자체교육 연 5회(의왕돌봄마당 4 + 재단 맞춤형 1)'],['찾아가는 통합돌봄 설명회','신규','통장협의회·동 협의체·경로당, 동별 연 1회(총 6회)'],['주민 생활 밀착 홍보','신규','경로당·약국·의원 홍보물, 퇴원 예정자 병원 안내문']];
[[ '운영기반 및 민관협력', ops, 0.6 ],[ '교육 및 홍보', edu, 6.75 ]].forEach(g=>{
  txt(s,g[0],{ x:g[2],y:1.4,w:5.95,h:0.5, fontSize:20, bold:true, color:C.text2 });
  g[1].forEach((o,i)=>{ const y=2.0+i*1.6; box(s,g[2],y,5.95,1.4,C.background2,'o'+g[2]+i);
    txt(s,o[1],{ x:g[2]+0.25,y:y+0.2,w:0.8,h:0.38, fontSize:12, bold:true, color:C.background1, fill:{ color: o[1]=='신규'?C.accent2:C.text2 }, align:'center', valign:'middle', margin:0 });
    txt(s,o[0],{ x:g[2]+1.2,y:y+0.12,w:4.6,h:0.55, fontSize:17, bold:true, color:C.text2, valign:'middle' });
    txt(s,o[2],{ x:g[2]+0.25,y:y+0.75,w:5.5,h:0.6, fontSize:14 }); });
});
s.addNotes('서비스마다 달랐던 의뢰 양식을 통일하고, 건보공단 등 전문기관과의 협업을 정례화합니다. 교육은 의왕돌봄마당으로 연 5회, 주민 대상 찾아가는 설명회는 동별 1회 이상 운영합니다.');

// 17 KPI
s=S('CONTENT','성과지표');
table(s,[[head('지표'),head('정의'),head('’27 목표'),head('근거')],
 ['통합돌봄 대상자 발굴률','직권·자체발굴 ÷ 신청',{text:'20% 이상',options:{bold:true}},'정부합동평가'],
 ['우선발굴필요군 중 대상자','6,506명 중 대상자 결정 인원',{text:'421명(6%)',options:{bold:true}},'수립 안내서'],
 ['대상자 중 우선발굴필요군 비중','우선발굴필요군 ÷ 전체 대상자',{text:'70% 이상',options:{bold:true}},'수립 안내서'],
 ['개인별 지원계획 수립률','수립자 ÷ 우선발굴필요군',{text:'6% 이상',options:{bold:true}},'지역복지사업 평가'],
 ['모니터링 수행률','3개월 단위 수행 ÷ 전체 대상자',{text:'80% 이상',options:{bold:true}},'수립 안내서'],
 ['전담인력 실제 배치율','배치 인원 ÷ 기준인건비 인원',{text:'100%',options:{bold:true}},'정부합동평가'],
 ['보건소 통합지원회의 참석률','보건소 참석 회의 ÷ 전체 회의',{text:'90% 이상',options:{bold:true}},'정부합동평가'],
 ['보건소 통합돌봄 서비스 이용률','보건소 2가지 이상 연계 건수 ÷ 연계 건수',{text:'25% 이상',options:{bold:true}},'정부합동평가']],
 { x:0.6,y:1.4,w:12.1,colW:[3.6,4.6,1.7,2.2], rowH:0.55, fontSize:14 });
txt(s,'자체 지표: 계획 변경·종결 반영 90%, 교육 이수 100%(전담 10명), 자체교육 연 5회, 설명회 연 6회, 만족도 조사 연 300명, 통합지원회의 월 2회 이상',{ x:0.6,y:6.5,w:12.1,h:0.45, fontSize:12, color:C.accent6 });
s.addNotes('성과지표는 행정안전부 2028년 지자체 합동평가와 보건복지부 지역복지사업 평가 기준에 맞췄습니다. 핵심은 발굴률 20% 이상입니다.');

// 18 budget
s=S('CONTENT','2027년 예산(안) 706,636천원');
s.addChart(pres.charts.BAR,[{ name:'예산', labels:['의왕살핌(일상·건강·내집)','신규서비스(이동·회복·AI)','의왕형 통합돌봄(보조사업)','운영비(회의·홍보·행사)','재택의료센터 운영지원','퇴원환자 연계수당'], values:[514136,110000,60000,17000,5000,500] }],
  { x:0.6,y:1.4,w:7.6,h:5.3, barDir:'bar', chartColors:[HX.navy], showValue:true, dataLabelPosition:'outEnd', dataLabelFormatCode:'#,##0', dataLabelFontSize:12, dataLabelColor:HX.text, catAxisLabelColor:HX.text, catAxisLabelFontSize:13, valAxisHidden:true, valGridLine:{ style:'none' }, catGridLine:{ style:'none' }, showLegend:false, catAxisOrientation:'maxMin', showTitle:true, title:'사업별 예산(천원)', titleFontSize:14, titleColor:HX.navy, catAxisLabelFontFace:'+mn-lt', dataLabelFontFace:'+mn-lt', titleFontFace:'+mn-lt' });
stat(s,8.6,1.5,4.1,2.5,'706,636','총 예산(천원)','사업비 689,636 + 운영비 17,000',true);
table(s,[[head('재원'),head('금액(천원)'),head('비율')],['국비','353,318','50%'],['도비','35,332','5%'],['시비','317,986','45%']],
 { x:8.6,y:4.3,w:4.1,colW:[1.2,1.8,1.1], rowH:0.5, align:'center' });
txt(s,'시의회 예산 심의 결과에 따라 변동될 수 있음',{ x:8.6,y:6.4,w:4.1,h:0.4, fontSize:11, color:C.accent6 });
s.addNotes('2027년 예산(안)은 706,636천원으로 국비 50%, 도비 5%, 시비 45%입니다. 의왕살핌에 514,136천원, 신규서비스 세 가지에 110,000천원을 편성했습니다. 시의회 심의 결과에 따라 달라질 수 있습니다.');

// 19 schedule & ask
s=S('CONTENT','향후 일정과 심의 요청');
const tl=[['11. 26.','협의체 심의'],['11. 30.','계획 확정·제출(경기도)'],['’27. 1.','동 전담인력 교육, 매뉴얼 개정'],['’27. 1~2월','제공기관 공모·협약'],['’27. 연중','초고령 조사, 설명회, 회의 월 2회']];
s.addShape(pres.shapes.LINE,{ x:0.9,y:2.15,w:11.5,h:0, line:{ color:HX.line, width:2 } });
tl.forEach((t,i)=>{ const x=0.6+i*2.45; s.addShape(pres.shapes.OVAL,{ x:x+0.95,y:1.95,w:0.4,h:0.4, fill:{ color: i==0?C.accent2:C.text2 }, line:{ color: i==0?C.accent2:C.text2 } });
  txt(s,t[0],{ x,y:2.5,w:2.3,h:0.45, fontSize:17, bold:true, color:C.text2, align:'center' }); txt(s,t[1],{ x,y:2.95,w:2.3,h:0.9, fontSize:14, align:'center' }); });
box(s,0.6,4.3,12.1,2.4,C.text2,'ask');
txt(s,'심의 요청',{ x:0.95,y:4.5,w:4,h:0.5, fontSize:20, bold:true, color:C.accent2 });
bullets(s,['2027년도 통합지원 지역계획(안)의 목표·핵심사업·성과지표의 적정성','대상자 발굴, 퇴원환자 연계, 방문의료 이용 확대를 위한 기관별 협력 방안'],{ x:0.95,y:5.1,w:11.4,h:1.4, fontSize:18, color:C.background1 });
s.addNotes('심의 결과는 11월 30일 계획 확정에 반영해 경기도에 제출하겠습니다. 계획의 적정성과 함께, 각 기관에서 발굴·퇴원환자 연계·방문의료 확대에 협력해 주실 부분을 말씀해 주시면 감사하겠습니다.');

(async()=>{ await pres.writeFile({ fileName:'deck.pptx' }); await applyTheme('deck.pptx', THEME); console.log('done'); })();
