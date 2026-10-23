'use strict';
// Five requested opening slides, built from the current DOCX and screenshot style.
const pptxgen = require('pptxgenjs');
const path = require('node:path');
const fs = require('node:fs');
const D = require('./diagrams');
const ROOT = path.resolve(__dirname, '../..');
const OUT = path.join(ROOT, 'downloads/presentation_5slides');
fs.mkdirSync(OUT, { recursive: true });
const pptx = new pptxgen();
pptx.defineLayout({name:'ACADEMIC',width:13.333333,height:7.5});
pptx.layout = 'ACADEMIC';
pptx.author = 'Research Presentation';
pptx.title = '太赫兹器件神经算子电磁建模：研究背景、现状与目标';
pptx.subject = '按用户指定范围制作的五页开篇；不含后续研究内容页';
pptx.company = '广州大学';
pptx.lang = 'zh-CN';
pptx.theme = {headFontFace:'Noto Sans CJK SC',bodyFontFace:'Noto Sans CJK SC',lang:'zh-CN'};
const C={blue:'155EA7',navy:'143B63',cyan:'2296BC',red:'BF2429',ink:'152332',gray:'536B80',
  border:'B9CDDD',pale:'F3F7FB',soft:'EAF2F9',white:'FFFFFF',gold:'D79B3A'};
const FONT='Noto Sans CJK SC';
const ST=pptx.ShapeType;
function text(s,t,x,y,w,h,size=18,color=C.ink,bold=false,align='left',extra={}) {
  s.addText(t,{x,y,w,h,fontFace:FONT,fontSize:size,color,bold,align,
    margin:0,breakLine:false,valign:'mid',fit:'none',paraSpaceAfterPt:0,
    objectName:typeof t==='string'?t:'原生富文本',...extra});
}
function rect(s,x,y,w,h,fill=C.white,line=C.border,lw=.8,extra={}) {
  s.addShape(ST.rect,{x,y,w,h,fill:{color:fill},line:{color:line,width:lw},...extra});
}
function line(s,x1,y1,x2,y2,color=C.border,width=.8,extra={}) {
  s.addShape(ST.line,{x:Math.min(x1,x2),y:Math.min(y1,y2),w:Math.max(.001,Math.abs(x2-x1)),h:Math.max(.001,Math.abs(y2-y1)),
    flipV:(x2-x1)*(y2-y1)<0,line:{color,width,...extra}});
}
function polygon(s,pts,fill,lineColor=fill,lw=0) {
  const x=Math.min(...pts.map(p=>p[0])),y=Math.min(...pts.map(p=>p[1]));
  const w=Math.max(...pts.map(p=>p[0]))-x,h=Math.max(...pts.map(p=>p[1]))-y;
  s.addShape(ST.custGeom,{x,y,w,h,points:[...pts.map(p=>({x:p[0]-x,y:p[1]-y})),{close:true}],
    fill:{color:fill},line:{color:lineColor,width:lw}});
}
function ribbon(s,label,x,y,w=1.65,h=.40) {
  polygon(s,[[x,y],[x+w,y],[x+w-.12,y+h],[x,y+h]],C.blue);
  polygon(s,[[x,y+h],[x+.14,y+h],[x+.14,y+h+.10]],C.navy);
  text(s,label,x+.12,y+.015,w-.28,h-.03,20,C.white,true,'center');
}
function arrow(s,x1,y1,x2,y2,color=C.cyan,width=1.7) {
  line(s,x1,y1,x2,y2,color,width,{endArrowType:'triangle'});
}
function picture(s,relative,x,y,w,h,pw,ph,alt) {
  const k=Math.min(w/pw,h/ph),iw=pw*k,ih=ph*k;
  s.addImage({path:path.join(ROOT,relative),x:x+(w-iw)/2,y:y+(h-ih)/2,w:iw,h:ih,
    altText:alt,objectName:alt});
}
function header(title,num,source) {
  const s=pptx.addSlide();s.background={color:C.white};
  text(s,title,.43,.16,10.85,.55,29,C.blue,true);
  text(s,'广州大学',11.25,.23,1.63,.38,18,C.blue,true,'right');
  line(s,0,.82,13.333333,.82,C.blue,1.65);
  // Thin extension follows the screenshot's academic blue-white frame.
  line(s,.43,7.13,12.90,7.13,C.border,.6);
  text(s,`资料依据：${source}`,.45,7.22,11.5,.16,8.5,C.gray);
  text(s,String(num),12.30,7.18,.55,.25,11,C.blue,false,'right');
  return s;
}
function banner(s,label,main,y=1.04,opts={}) {
  rect(s,.53,y+.20,12.36,.68,'FAFCFE',C.border,.9);
  ribbon(s,label,.40,y,1.67,.40);
  text(s,main,.76,y+.36,11.80,.35,opts.size||21,C.navy,true,'center');
}
function boxTitle(s,t,x,y,w) { rect(s,x,y,w,.43,C.soft,C.border,.7);text(s,t,x+.10,y+.035,w-.20,.35,19,C.blue,true,'center'); }
function summary(s,t,y=6.70,size=19.5) {
  rect(s,.43,y,12.47,.36,'EEF4FA','EEF4FA',0);
  text(s,t,.60,y+.025,12.12,.30,size,C.blue,true,'center');
}
function sparkPlot(s,x,y,w,h) {
  line(s,x+.10*w,y+.08*h,x+.10*w,y+.90*h,C.gray,.7);
  line(s,x+.10*w,y+.90*h,x+.98*w,y+.90*h,C.gray,.7);
  let prev;
  for(let i=0;i<=55;i++) {
    const u=i/55,v=.23+.57*Math.exp(-Math.pow((u-.39)/.065,2))+.38*Math.exp(-Math.pow((u-.74)/.08,2));
    const p=[x+(.10+.88*u)*w,y+v*h];
    if(prev)line(s,...prev,...p,C.blue,1.25);prev=p;
  }
}
function note(s,t,source) {
  s.addNotes(`本页范围：${source}\n内容依据本轮上传《太赫兹国重申请书-汇总-10-22.docx》，SHA256: 15d8ce5a779f0d1df8fd64349a37fbcf44769be8e36d077b24ab54c103c6b70c。\n${t}\n版式参考用户提供的“1.1立项依据与研究现状”截图。上传《广州大学-张珍TZ4.pptx》超过当前文件传输工具32MiB限制，未读取或套用其母版。页眉“广州大学”为文字标识，并非新建或冒用官方校徽。\n图形为原生可编辑概念示意，应用场景资产复用此前生成的独立科研概念图，不是实测或仿真结果。`);
}

// 1. Background, domain status and R&D tasks; no method taxonomy or critique.
{
 const s=header('1.1 研究背景及现状',1,'需求分析；研究现状首段');
 banner(s,'研究背景',[
   {text:'太赫兹应用牵引 ',options:{color:C.navy}},
   {text:'高性能、多功能器件',options:{color:C.red}},
   {text:' 研发，电磁场建模支撑 ',options:{color:C.navy}},
   {text:'机理分析与结构设计',options:{color:C.red}}
 ],1.02,{size:20.5});
 rect(s,.43,2.10,12.47,2.33,'F7FAFD',C.border,.8);
 boxTitle(s,'应用牵引',.59,2.23,3.10);
 const apps=[['communication','高速通信'],['security','安检成像'],['biosensing','生物传感'],['material_identification','物质识别']];
 apps.forEach(([file,label],i)=>{
   const x=.65+(i%2)*1.46,y=2.79+Math.floor(i/2)*.70;
   picture(s,`assets/figure1_v2/${file}.png`,x,y,1.22,.45,1536,1024,`${label}应用概念图`);
   text(s,label,x-.05,y+.43,1.32,.27,15.5,C.navy,true,'center');
 });
 arrow(s,3.77,3.34,4.15,3.34);
 text(s,'亚波长太赫兹器件',4.17,2.27,3.80,.35,20.5,C.navy,true,'center');
 picture(s,'assets/figure1_v3/core_device.png',4.55,2.65,3.10,1.33,1448,1086,'薄膜与狭缝构成的亚波长器件概念图');
 text(s,'薄膜 · 狭缝 · 复杂边界',4.20,4.02,3.80,.30,17,C.red,true,'center');
 arrow(s,8.01,3.34,8.43,3.34);
 text(s,'三维场决定器件响应',8.49,2.27,4.12,.36,20.5,C.navy,true,'center');
 D.fieldCube(s,pptx,8.95,2.85,1.35,1.04);
 sparkPlot(s,10.79,2.84,1.37,.85);
 arrow(s,10.37,3.35,10.66,3.35,C.cyan,1.3);
 text(s,'传播 · 耦合 · 谐振 · 优化',8.57,4.02,3.95,.30,17,C.navy,false,'center');
 ribbon(s,'领域现状',.43,4.63,1.65,.40);
 ribbon(s,'研发任务',6.77,4.63,1.65,.40);
 rect(s,.56,5.11,5.91,1.24,C.pale,C.border,.8);
 text(s,'国内外围绕太赫兹器件电磁建模\n已开展大量研究工作',.76,5.20,5.51,.58,19,C.ink,false,'center');
 text(s,'关注器件电磁场分布及其电磁响应',.76,5.96,5.51,.27,17,C.blue,true,'center');
 rect(s,6.90,5.11,5.87,1.24,C.pale,C.border,.8);
 text(s,'获取不同构型下的三维电磁场\n开展机理分析、性能评估与结构优化',7.09,5.20,5.49,.58,18.7,C.ink,false,'center');
 text(s,'面向器件快速设计与自主建模工具',7.09,5.96,5.49,.27,17,C.blue,true,'center');
 summary(s,'以三维电磁场联系器件结构与电磁性能，服务太赫兹器件研发。',6.70,19.5);
 note(s,'本页仅讲应用背景、领域研究积累与研发任务，不出现方法分类、优劣或挑战。[0048]支撑应用、精细器件、三维场与性能/结构设计关系；[0055]首段说明国内外已有大量建模研究；[0052]支撑快速设计与自主工具的需求导向，不表述为已经实现的领域能力。','[0048]、[0052]、[0055]');
}

// 2. Three method families, compared by the quantity that is predicted.
{
 const s=header('1.2 现有电磁建模方法',2,'研究现状2.1、2.2、2.3');
 text(s,'从数值求解到数据驱动：关键区别在于“预测什么、如何求解”',.55,1.05,12.20,.42,21,C.navy,true,'center');
 const xs=[.43,4.65,8.87],w=4.03;
 const titles=['传统全波仿真','电磁响应建模','电磁场建模'];
 const methods=['FDTD · FEM · MoM · FIT','正向预测 · 逆向设计 · 物理先验','FNO · DeepONet · 几何增强算子'];
 xs.forEach((x,i)=>{
   rect(s,x,1.70,w,4.85,C.pale,C.border,.9);
   rect(s,x,1.70,w,.55,C.blue,C.blue,0);
   text(s,`${['2.1','2.2','2.3'][i]}  ${titles[i]}`,x+.13,1.80,w-.26,.32,21,C.white,true,'center');
   text(s,methods[i],x+.12,2.36,w-.24,.32,15.7,C.blue,true,'center');
 });
 text(s,'几何 / 材料 / 激励',.69,2.91,3.50,.35,18,C.navy,true,'center');
 D.meshCube(s,pptx,1.64,3.39,1.50,.94);
 text(s,'网格离散与数值求解',.72,4.40,3.44,.32,18,C.ink,false,'center');
 text(s,'E、H场 → 电磁响应',.71,4.84,3.46,.33,18,C.blue,true,'center');

 text(s,'设计变量 ⇄ 学习模型',4.90,2.91,3.52,.35,18,C.navy,true,'center');
 D.network(s,pptx,4.98,3.46,1.48,.80);sparkPlot(s,6.99,3.44,1.15,.77);
 arrow(s,6.60,3.86,6.89,3.86);
 text(s,'正向预测 / 逆向设计',4.95,4.40,3.42,.32,18,C.ink,false,'center');
 text(s,'S参数 · 光谱 · 方向图',4.93,4.84,3.46,.33,18,C.blue,true,'center');

 text(s,'几何表征 / 激励条件',9.11,2.91,3.54,.35,18,C.navy,true,'center');
 rect(s,9.23,3.49,1.36,.70,'E5F0F9',C.blue,1);
 text(s,'神经算子',9.28,3.68,1.26,.29,17,C.blue,true,'center');
 arrow(s,10.69,3.84,10.97,3.84);D.fieldCube(s,pptx,11.05,3.38,1.18,.95);
 text(s,'学习输入条件到全场的映射',9.10,4.40,3.57,.32,17.4,C.ink,false,'center');
 text(s,'三维 E/H → 电磁响应',9.12,4.84,3.52,.33,18,C.blue,true,'center');

 const basis=['建模基础：麦克斯韦方程','正向：结构参数 → 电磁响应','学习对象：偏微分方程解算子'];
 const procedure=['通过网格离散形成数值方程\n求解电磁场后提取器件响应','逆向：目标响应 → 结构参数\n基于样本建立设计变量与响应映射','建立几何 / 激励到全场的映射\n由预测场进一步提取器件响应'];
 xs.forEach((x,i)=>{
   line(s,x+.17,5.31,x+w-.17,5.31,C.border,.7);
   text(s,basis[i],x+.17,5.44,w-.34,.34,16.6,C.blue,true,'center');
   text(s,procedure[i],x+.18,5.93,w-.36,.49,16.5,C.ink,false,'center');
 });
 summary(s,'三类方法分别通过数值求解、响应映射与全场算子学习完成电磁建模。',6.73,18.9);
 note(s,'本页只介绍2.1、2.2、2.3的方法原理、代表形式、输入与输出，不列优势、弊端或未来挑战。原文参考文献分组：[3]–[5]对应全波仿真，[6]–[17]对应电磁响应建模，[18]–[30]对应神经算子及其变体。没有使用文献性能数字。','[0059]–[0070]');
}

// 3. Literature limitations and the resulting open scientific problems.
{
 const s=header('1.3 现有研究面临的具体挑战',3,'2.3末段及2.4研究领域存在的挑战');
 text(s,'面向亚波长太赫兹器件，亟待解决三个关键挑战',.55,1.06,12.20,.43,22,C.navy,true,'center');
 text(s,'关键挑战',.53,1.67,2.18,.30,16,C.blue,true,'center');
 text(s,'复杂性来源',3.08,1.67,5.67,.30,16,C.blue,true,'center');
 text(s,'需要回答的问题',9.13,1.67,3.64,.30,16,C.blue,true,'center');
 const rows=[
   {y:2.10,title:'几何信息完整性',lead:'尺度、方向与亚网格细节需协同保留',body:'零厚表面、物理棱边与角点交汇关系复杂\n有限规则网格需承载精细边界信息',q:'如何保留精细边界\n及其局部几何信息？'},
   {y:3.60,title:'三维场预测',lead:'局部边界变化与物理约束共同作用',body:'局部边界方向变化影响三维场分布\n几何测度、算子表达与物理约束需匹配',q:'如何融合几何与物理约束\n实现三维场的局部增强？'},
   {y:5.10,title:'跨任务泛化',lead:'新频段与新拓扑下的知识复用',body:'跨任务几何与电磁响应存在差异\n需以少量目标样本支持多器件建模应用',q:'如何用少量样本实现\n跨频率、跨构型适配？'}
 ];
 rows.forEach((r,i)=>{
   rect(s,.43,r.y,12.47,1.34,i%2?'F7FAFD':C.pale,C.border,.8);
   text(s,r.title,.53,r.y+.08,2.18,.35,20,C.blue,true,'center');
   if(i===0){
     D.ambiguousBoundary(s,pptx,.84,r.y+.47,1.56,.56);
     text(s,'同面积 · 异朝向',.53,r.y+1.08,2.18,.19,13.4,C.gray,false,'center');
   }
   if(i===1)D.fieldCube(s,pptx,1.08,r.y+.49,1.08,.70);
   if(i===2)D.transfer(s,pptx,.78,r.y+.52,1.70,.64);
   text(s,r.lead,3.08,r.y+.12,5.70,.35,18.5,C.navy,true);
   text(s,r.body,3.08,r.y+.58,5.70,.62,17.5,C.ink);
   rect(s,9.13,r.y+.16,3.50,1.00,'FFF7F6','E3BFC1',.8);
   text(s,r.q,9.29,r.y+.30,3.18,.70,19,C.red,true,'center');
 });
 summary(s,'三项挑战分别指向精细几何信息保留、三维电磁场预测与跨任务泛化。',6.73,18.5);
 note(s,'本页仅集中呈现2.4及其前文的三个科学挑战，不重复2.1—2.3的方法分类和优劣对照。第一项对应[0069]、[0074]的尺度/方向/位置与边界交汇信息；第二项对应[0070]、[0075]的几何测度与算子匹配、麦克斯韦物理约束及三维场；第三项对应[0071]、[0076]的跨频率/跨构型、小样本适配和多器件应用。第5页仍保留本项目实现接口的技术难点。','[0069]–[0076]');
}

// 4. Proposed capabilities, strictly from the research-objective section.
{
 const s=header('2.1 研究目标',4,'二、研究目标');
 banner(s,'总体目标','高保真几何输入 → 三维复电磁场预测 → 新频段 / 新构型适配',1.02,{size:21.5});
 const xs=[.50,4.72,8.94],w=3.89;
 const goals=[
  ['目标一','精细边界高保真表征','提出分维几何测度表征方法\n同时编码尺度、方向与位置','形成高保真几何输入'],
  ['目标二','三维复电磁场快速预测','构建空间变系数傅里叶神经算子\n融合麦克斯韦方程物理约束','建立几何 / 激励到 E/H 的映射'],
  ['目标三','跨频率与跨构型适配','提出迁移学习与数据增强方法\n建立跨任务知识迁移机制','提升新器件、新频段复用能力']
 ];
 goals.forEach((g,i)=>{
  const x=xs[i];rect(s,x,2.12,w,3.72,C.pale,C.border,.8);
  rect(s,x,2.12,w,.46,C.blue,C.blue,0);text(s,g[0],x+.10,2.19,w-.20,.31,21,C.white,true,'center');
  text(s,g[1],x+.12,2.75,w-.24,.39,20,C.navy,true,'center');
  if(i===0)D.geometryMeasures(s,pptx,x+.51,3.38,2.87,.83);
  if(i===1)D.fieldCube(s,pptx,x+1.25,3.28,1.38,1.02);
  if(i===2)D.transfer(s,pptx,x+.57,3.40,2.75,.79);
  text(s,g[2],x+.15,4.46,w-.30,.65,17.5,C.ink,false,'center');
  rect(s,x+.17,5.29,w-.34,.37,C.white,C.border,.7);
  text(s,g[3],x+.24,5.33,w-.48,.28,i===1?16.2:16.8,C.red,true,'center');
 });
 xs.forEach(x=>arrow(s,x+w/2,5.87,x+w/2,6.10,C.cyan,1.7));
 line(s,2.45,6.10,10.89,6.10,C.cyan,1.1);
 arrow(s,6.67,6.10,6.67,6.29,C.cyan,1.5);
 rect(s,2.08,6.33,9.17,.51,C.blue,C.blue,0);
 text(s,'支撑自主智能电磁建模与太赫兹器件高效设计',2.28,6.42,8.77,.31,21.5,C.white,true,'center');
 note(s,'三项目标完全依“二、研究目标”进行能力化概括，不加入第四章定量考核指标，不表述为已经实现的成果。整体自主工具与高效设计支撑参见需求分析[0052]及工具集成[0296]。','[0112]–[0117]；总体支撑语依据[0052]、[0296]');
}

// 5. The project's implementation interfaces, handing off to existing content slides.
{
 const s=header('2.2 研究目标实现面临的困难',5,'二、研究目标；三、研究内容');
 text(s,'从目标到可用模型，需要打通三个关键接口',.55,1.05,12.20,.44,22,C.navy,true,'center');
 text(s,'研究目标',.49,1.64,2.30,.31,16,C.blue,true,'center');
 text(s,'实现难点',3.28,1.64,6.20,.31,16,C.blue,true,'center');
 text(s,'后续研究内容入口',10.26,1.64,2.50,.31,16,C.blue,true,'center');
 const rows=[
  ['高保真\n几何输入','真实物理边界提取与统一编码','区分真实棱边 / 角点与剖分辅助要素\n统一多维测度、方向与亚网格位置信息','研究内容一\n精细边界表征'],
  ['快速准确\n场预测','局部几何、求解效率与物理约束的协同','边界朝向变化引起局部场突变\n变系数计算、FFT效率与物理约束需兼顾','研究内容二\n空间变系数FNO'],
  ['新任务\n快速适配','少样本下的任务对齐与物理一致迁移','频率伸缩受色散、电导率及耦合变化影响\n跨拓扑难对齐，增强样本需保持物理一致','研究内容三\n迁移与数据增强']
 ];
 rows.forEach((r,i)=>{
   const y=2.10+i*1.48;rect(s,.43,y,12.47,1.31,C.pale,C.border,.8);
   rect(s,.54,y+.13,2.19,1.05,C.blue,C.blue,0);
   text(s,r[0],.66,y+.31,1.95,.66,21,C.white,true,'center');
   arrow(s,2.82,y+.66,3.11,y+.66,C.cyan,1.7);
   text(s,r[1],3.26,y+.12,6.45,.37,19.5,C.red,true);
   text(s,r[2],3.26,y+.61,6.45,.56,17.4,C.ink);
   arrow(s,9.77,y+.66,10.04,y+.66,C.cyan,1.7);
   rect(s,10.16,y+.17,2.49,.97,C.white,C.border,.9);
   text(s,r[3],10.29,y+.30,2.23,.71,18.5,C.blue,true,'center');
 });
 summary(s,'下文分别围绕三个关键接口，展开研究内容与解决思路。',6.73,20);
 note(s,'本页是承上启下的目标实现难点，不重复第3页的研究现状，也不代做用户已有的研究内容部分。第一行来源于真实物理边界提取及统一几何编码；第二行来源于几何方向、变系数算子计算开销、FFT效率及物理约束协同；第三行来源于色散/电导率/耦合导致频率伸缩偏差、不同拓扑/端口构型对齐与物理一致的增强样本。未展开23通道、六方向分解、MAML内外循环等技术细节。','[0126]、[0144]、[0174]、[0178]、[0199]–[0220]、[0256]、[0262]–[0269]、[0274]–[0307]');
}

pptx.writeFile({fileName:path.join(OUT,'thz_background_objectives_5slides_revised.pptx')})
  .then(()=>console.log('Built exactly five slides:',OUT))
  .catch(e=>{console.error(e);process.exitCode=1;});
