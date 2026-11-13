import {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,AlignmentType,HeadingLevel,ExternalHyperlink,InternalHyperlink,Bookmark,ImageRun,BorderStyle,PageNumber,Footer,Header,ShadingType,TableLayoutType} from 'docx';
import {readFile} from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import {CATALOG,MODELS,SPECIAL} from './_catalog.mjs';
const BLUE='007CF3',GREEN='08753B',INK='10283B',SLATE='5D7182',LIGHT='EAF4FC',PALE='F1F9F5',WHITE='FFFFFF';
const clean=(v,max=300)=>typeof v==='string'?v.trim().slice(0,max):'';
const safeUrl=value=>{try {const u=new URL(value); return ['https:','http:'].includes(u.protocol)&&u.username===''&&u.password===''&&u.href.length<=1800?u.href:null;}catch{return null;}};
const entitySafe=s=>String(s||'').replace(/[\u0000-\u0008\u000b-\u001f]/g,'');
export function validate(body){
 if(!body||typeof body!=='object'||Array.isArray(body))throw new Error('Enter your company details.');
 const company=clean(body.company,100),industry=clean(body.industry,100),model=clean(body.model,70);
 if(company.length<2||industry.length<2)throw new Error('Enter your company name and industry.');
 if(!MODELS.includes(model))throw new Error('Choose a valid business model.');
 const teamSize=Number(body.teamSize||0);if(!Number.isInteger(teamSize)||teamSize<0||teamSize>100000)throw new Error('Enter a valid team size.');
 const horizon=clean(body.horizon,35);if(!['0–30 days','31–90 days','91–180 days','6–12 months','Longer-term / exploring'].includes(horizon))throw new Error('Choose your exit timeframe.');
 const raw=clean(body.links,14000);let links=[];
 if(raw){const lines=raw.split(/\r?\n/).filter(x=>x.trim());if(lines.length>65)throw new Error('Use up to 65 existing document links.');
 for(let i=0;i<lines.length;i++){
 const p=lines[i].split('|').map(x=>x.trim()); if(p.length<2||p.length>3||!p[0]||!p[1]||p[0].length>120||p[1].length>1800)throw new Error(`Fix existing document link on line ${i+1}: Title | https://... | Department (optional).`);
 const url=safeUrl(p[1]);if(!url)throw new Error(`Use a valid public-style HTTPS/HTTP link on line ${i+1}.`);
 links.push({title:p[0],url,department:p[2]||''});
 }}
 return {company,industry,model,teamSize,horizon,owner:clean(body.owner,100),description:clean(body.description,650),priority:clean(body.priority,300),systems:clean(body.systems,450),departments:clean(body.departments,300),links,prepared:new Date().toISOString().slice(0,10)};
}
export function buildInventory(input){
 const categories=Object.entries(CATALOG).map(([name,titles])=>({name,entries:titles.map(title=>({title,url:null,source:'Suggested / not provided'}))}));
 const operations=categories.find(x=>x.name==='Operations & Fulfillment');
 operations.entries.push(...SPECIAL[input.model].map(title=>({title,url:null,source:'Model-specific / not provided'})));
 const remaining=[];
 for(const link of input.links){
  const lower=link.title.toLowerCase();const target=categories.find(x=>x.name.toLowerCase()===link.department.toLowerCase()) || categories.find(x=>x.entries.some(e=>e.title.toLowerCase()===lower));
  if(target){const match=target.entries.find(e=>e.title.toLowerCase()===lower);
    if(match){match.url=link.url;match.source='Member supplied';}else target.entries.push({title:link.title,url:link.url,source:'Member supplied'});
  }else remaining.push({title:link.title,url:link.url,source:'Member supplied'});
 }
 if(remaining.length)categories.push({name:'Other Existing Documents',entries:remaining});
 return categories;
}
const para=(text,options={})=>new Paragraph({children:[new TextRun({text:entitySafe(text),...options.run})],spacing:{after:120},...options});
const linkRun=(label,url)=>new ExternalHyperlink({link:url,children:[new TextRun({text:label,color:BLUE,underline:{color:BLUE},bold:true,size:18})]});
const cell=(content,fill=WHITE)=>new TableCell({shading:{fill},margins:{top:105,bottom:95,left:125,right:100},children:[typeof content==='string'?para(content,{spacing:{after:0},run:{size:18}}):new Paragraph({children:content,spacing:{after:0}})]});
const row=(a,b,c,d,index)=>new TableRow({children:[cell(a,index%2?WHITE:'F5F9FC'),cell(b,index%2?WHITE:'F5F9FC'),cell(c,index%2?WHITE:'F5F9FC'),cell(d,index%2?WHITE:'F5F9FC')]});
const label=(k,v)=>new Paragraph({children:[new TextRun({text:k+'  ',bold:true,color:INK,size:20}),new TextRun({text:v||'Not provided',color:v?SLATE:'A65A23',size:20})],spacing:{after:125}});
const rules={top:{style:BorderStyle.SINGLE,size:3,color:'DDEAF2'},bottom:{style:BorderStyle.SINGLE,size:3,color:'DDEAF2'},left:{style:BorderStyle.SINGLE,size:3,color:'DDEAF2'},right:{style:BorderStyle.SINGLE,size:3,color:'DDEAF2'},insideHorizontal:{style:BorderStyle.SINGLE,size:3,color:'DDEAF2'},insideVertical:{style:BorderStyle.SINGLE,size:3,color:'DDEAF2'}};
function inventoryTable(entries){
 const header=new TableRow({tableHeader:true,children:['Document / asset','Status','Link / location','Next action / owner'].map(t=>new TableCell({shading:{fill:INK},margins:{top:125,bottom:125,left:125,right:125},children:[para(t,{spacing:{after:0},run:{color:WHITE,bold:true,size:17}})]}))});
 const lines=entries.map((e,i)=>row(e.title,e.url?'LINK PROVIDED':'TO ADD',e.url?[linkRun('Open document',e.url)]:[new TextRun({text:'Paste link here: __________________',color:'8295A6',size:17})],e.url?'Verify permissions / assign document owner':'Locate or create; assign owner',i));
 return new Table({width:{size:100,type:WidthType.PERCENTAGE},layout:TableLayoutType.FIXED,borders:rules,rows:[header,...lines],columnWidths:[3500,1350,2400,2750]});
}
export async function createMaster(input){
 const inventory=buildInventory(input);const total=inventory.reduce((n,c)=>n+c.entries.length,0);const linked=inventory.reduce((n,c)=>n+c.entries.filter(e=>e.url).length,0);
 const children=[];
 const img=await readFile(fileURLToPath(new URL('../../public/cashout-challenge-logo.png',import.meta.url)));
 children.push(new Paragraph({children:[new ImageRun({data:img,transformation:{width:285,height:101},type:'png'})],spacing:{after:250}}));
 children.push(new Paragraph({children:[new TextRun({text:'COMPANY MASTER',bold:true,color:INK,size:54})],spacing:{before:200,after:100}}));
 children.push(new Paragraph({children:[new TextRun({text:'The operating file behind your business',color:BLUE,size:23,bold:true})],spacing:{after:250}}));
 children.push(new Paragraph({children:[new TextRun({text:input.company,bold:true,color:INK,size:34})],spacing:{after:140}}));
 children.push(para('Central document directory  •  Ownership map  •  Deal-readiness register',{run:{color:SLATE,size:20}}));
 children.push(new Paragraph({children:[new TextRun({text:`${linked} linked  /  ${total} identified  /  ${total-linked} awaiting links`,bold:true,color:GREEN,size:22})],shading:{fill:PALE},spacing:{before:180,after:230}}));
 children.push(label('Industry',input.industry));children.push(label('Business model',input.model));children.push(label('Team size',String(input.teamSize||'Not provided')));children.push(label('Exit horizon',input.horizon));children.push(label('Prepared',input.prepared));children.push(label('Master document owner',input.owner));
 if(input.description)children.push(label('Business description',input.description));if(input.priority)children.push(label('Most important documentation gap',input.priority));
 children.push(new Paragraph({text:'START HERE',heading:HeadingLevel.HEADING_1,pageBreakBefore:true}));
 children.push(para('This is your editable company master index. It is not a repository of the underlying files. Keep your documents in your existing secure company storage, then replace or add clickable links here. A link only works for someone who already has permission to open its destination.'));
 children.push(para('How to use: (1) Open the relevant department below, (2) link the actual company-owned file or folder, (3) write its responsible owner and review cadence, (4) confirm access with a non-owner manager, and (5) revisit before sharing with any buyer.')); 
 children.push(para('Suggested records are not claims that the files exist. “TO ADD” means no existing link was supplied. A link marked “LINK PROVIDED” has not been independently checked or verified.'));
 children.push(new Paragraph({text:'DEPARTMENT INDEX',heading:HeadingLevel.HEADING_2}));
 inventory.forEach((group,i)=>children.push(new Paragraph({children:[new InternalHyperlink({anchor:`dept_${i}`,children:[new TextRun({text:`${String(i+1).padStart(2,'0')}   ${group.name}   (${group.entries.filter(x=>x.url).length}/${group.entries.length} linked)`,color:BLUE,underline:{color:BLUE},size:21})]})],spacing:{after:125}})));
 children.push(new Paragraph({text:'90-DAY FILE CLEANUP PLAN',heading:HeadingLevel.HEADING_2}));
 children.push(para('Week 1: Assign a document owner and confirm the systems where records actually live. Prioritize your financials, contracts, company org chart, KPI dashboard, and operational SOPs.'));
 children.push(para('Weeks 2–4: Collect and link the high-priority records; ensure the company owns the accounts and can grant access. Identify out-of-date or missing materials and assign a responsible person.'));
 children.push(para('Days 30–90: Finish the function-specific records, document key-person dependencies, validate financial/source records, and prepare a restricted buyer-access copy when appropriate.'));
 if(input.systems)children.push(label('Existing systems / file repositories',input.systems));
 if(input.departments)children.push(label('Priority departments',input.departments));
 for(let i=0;i<inventory.length;i++){
  const group=inventory[i];const done=group.entries.filter(x=>x.url).length;
  children.push(new Paragraph({children:[new Bookmark({id:`dept_${i}`,children:[new TextRun({text:`${String(i+1).padStart(2,'0')}  ${group.name.toUpperCase()}`,color:INK,bold:true,size:30})]})],pageBreakBefore:true,spacing:{after:120},keepNext:true}));
  children.push(para(`${done} of ${group.entries.length} linked  •  Suggested department owner: [Assign role]  •  Review cadence: [Monthly / quarterly]`,{run:{color:SLATE,size:18}}));
  children.push(inventoryTable(group.entries));
  children.push(para('Department notes / exceptions: ______________________________________________________________',{spacing:{before:150,after:80},run:{color:SLATE,size:17}}));
 }
 children.push(new Paragraph({text:'DOCUMENT CONTROL & HANDOFF',heading:HeadingLevel.HEADING_1,pageBreakBefore:true}));
 children.push(para('Master record owner: [Name / role]     Last verified: [Date]     Next review: [Date]'));
 children.push(para('Source of truth / storage system: [Company-controlled repository]'));
 children.push(para('Permission model: [Staff / leadership / legal / buyer-restricted]'));
 children.push(para('Before sale diligence: remove personal payroll, bank credentials, secret keys, and unnecessary customer personal data from any shared copy. Use a restricted data room and confirm professional advice for legal, tax, and transaction-sensitive materials.'));
 const doc=new Document({creator:'Cash Out Challenge',title:`${input.company} | Company Master`,description:'Customizable company master document and file directory',styles:{default:{document:{run:{font:'Aptos',size:20,color:INK},paragraph:{spacing:{after:100}}}},paragraphStyles:[{id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,run:{font:'Aptos Display',bold:true,color:INK,size:30},paragraph:{spacing:{before:240,after:140}}},{id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,run:{bold:true,color:BLUE,size:24},paragraph:{spacing:{before:240,after:120}}}]},sections:[{properties:{page:{size:{width:12240,height:15840},margin:{top:850,bottom:800,left:750,right:750}}},footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:'CASH OUT CHALLENGE  •  COMPANY MASTER  |  ',color:SLATE,size:14}),new TextRun({children:[PageNumber.CURRENT],color:SLATE,size:14})]})]})},children}]});
 return Buffer.from(await Packer.toBuffer(doc));
}
