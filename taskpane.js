/* Excel AI Assistant - Gemini API client. API key is session-only. */
let selectedText = "";
let lastAnswer = "";
let conversation = [];
const $ = id => document.getElementById(id);
function addMessage(text, who="assistant", isError=false) {
  const el=document.createElement("div");
  el.className="msg "+(who==="user"?"user":"assistant")+(isError?" error":"");
  el.textContent=text;
  $("chat").appendChild(el);
  $("chat").scrollTop=$("chat").scrollHeight;
}
function status(t){$("status").textContent=t;}
Office.onReady(info=>{
  if(info.host==="Excel"){
    $("readBtn").onclick=readSelection;
    $("sendBtn").onclick=sendPrompt;
    $("clearBtn").onclick=()=>{conversation=[];$("chat").innerHTML="";addMessage("Chat cleared. What would you like to do?");};
    $("insertBtn").onclick=insertAnswer;
    $("key").value=sessionStorage.getItem("geminiKey")||"";
    $("key").addEventListener("input",()=>sessionStorage.setItem("geminiKey",$("key").value.trim()));
  } else status("Open this add-in inside Excel.");
});
async function readSelection(){
  status("Reading selected range…");
  try{
    await Excel.run(async ctx=>{
      const range=ctx.workbook.getSelectedRange();
      range.load(["address","values"]);
      await ctx.sync();
      const values=range.values||[];
      const safe=values.slice(0,30).map(row=>row.slice(0,20));
      selectedText=safe.map(r=>r.map(v=>String(v??"")).join("\t")).join("\n");
      $("selection").textContent="Range: "+range.address+"\n"+selectedText;
      status("Selected cells loaded.");
    });
  }catch(e){status("Excel error: "+e.message);}
}
async function sendPrompt(){
  const prompt=$("prompt").value.trim();
  const key=$("key").value.trim();
  if(!prompt){status("Type a question first.");return;}
  if(!key){status("Please add your Gemini API key.");return;}
  const fullPrompt=(selectedText?"Selected Excel data:\n"+selectedText+"\n\n":"")+
    "User question: "+prompt+"\nRespond clearly. If writing an Excel formula, state which cell/range it applies to and use the user's Excel locale where possible.";
  addMessage(prompt,"user");
  $("prompt").value="";
  $("sendBtn").disabled=true;status("AI is thinking…");
  try{
    const history=conversation.slice(-8).map(x=>({role:x.role,parts:[{text:x.text}]}));
    history.push({role:"user",parts:[{text:fullPrompt}]});
    const model="gemini-2.5-flash";
    const url=`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(key)}`;
    const response=await fetch(url,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({contents:history,generationConfig:{temperature:0.4,maxOutputTokens:1200}})});
    const data=await response.json();
    if(!response.ok)throw new Error(data?.error?.message||("HTTP "+response.status));
    const answer=data?.candidates?.[0]?.content?.parts?.map(p=>p.text||"").join("")||"No response text.";
    addMessage(answer);
    lastAnswer=answer;
    conversation.push({role:"user",text:fullPrompt},{role:"model",text:answer});
    status("Done.");
  }catch(e){addMessage("Request failed: "+e.message,"assistant",true);status("Error. Check key, internet, and API quota.");}
  finally{$("sendBtn").disabled=false;}
}
async function insertAnswer(){
  if(!lastAnswer){status("No AI answer to insert yet.");return;}
  try{
    await Excel.run(async ctx=>{
      const range=ctx.workbook.getSelectedRange();
      range.values=[[lastAnswer]];
      await ctx.sync();
    });
    status("Answer inserted into selected cell.");
  }catch(e){status("Insert error: "+e.message);}
}