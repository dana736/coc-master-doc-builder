import {createHash} from 'node:crypto';
import {getStore} from '@netlify/blobs';
import {json,verifyTurnstile,clientIp} from './_shared.mjs';
import {validate,createMaster} from './_builder.mjs';
const MAX_IP_DAILY=Math.max(1,Math.min(30,Number(process.env.DAILY_PER_IP_LIMIT)||3));
const MAX_GLOBAL_DAILY=Math.max(1,Math.min(10000,Number(process.env.DAILY_GLOBAL_LIMIT)||200));
export default async request=>{
 if(request.method!=='POST')return json({error:'Method not allowed'},405);
 try{
  const length=Number(request.headers.get('content-length')||0);
  if(length>21000)return json({error:'Too many characters in this request.'},413);
  const raw=await request.text();if(raw.length>21000)return json({error:'Too many characters in this request.'},413);
  let body;try{body=JSON.parse(raw)}catch{return json({error:'Invalid form data.'},400)}
  if(!await verifyTurnstile(body.turnstileToken,request))return json({error:'Please complete the security check and try again.'},403);
  let input;try{input=validate(body)}catch(e){return json({error:e.message},400)}
  const store=getStore('cim-master-generation-limits');const today=new Date().toISOString().slice(0,10);
  const ipKey=`ip-${today}-${createHash('sha256').update(clientIp(request)).digest('hex')}`;const globalKey=`global-${today}`;
  const [ip,global]=await Promise.all([store.get(ipKey,{type:'json'}),store.get(globalKey,{type:'json'})]);
  if((ip?.count||0)>=MAX_IP_DAILY||(global?.count||0)>=MAX_GLOBAL_DAILY)return json({error:'Daily generation limit reached. Try again tomorrow or contact support.'},429);
  const file=await createMaster(input);
  await Promise.all([store.setJSON(ipKey,{count:(ip?.count||0)+1}),store.setJSON(globalKey,{count:(global?.count||0)+1})]).catch(()=>console.error('Rate counter update failed'));
  const safe=input.company.replace(/[^a-z0-9-_ ]/gi,'').trim().replace(/\s+/g,'-').slice(0,45)||'Business';
  return new Response(new Uint8Array(file),{headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','Content-Disposition':`attachment; filename="Cash-Out-Company-Master-${safe}.docx"`,'Cache-Control':'no-store','X-Content-Type-Options':'nosniff'}});
 }catch(e){console.error('CIM generation failed',e?.name);return json({error:'Unable to build your Master Document. Check your entries and try again.'},502)}
};
