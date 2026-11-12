// RKL Approval Web App — paste into a new Apps Script file (ApprovalWebApp.gs).
const RKL_SHEET_ID = '1qVU9dsdxm3HfgQT0LS5uee3NgaC0igihY_Jh2zvq7q8';
function doGet(e) {
  const post = String((e && e.parameter && e.parameter.post) || '');
  const html = HtmlService.createHtmlOutput(`<!doctype html><html><head><base target="_top"><meta name="viewport" content="width=device-width,initial-scale=1"><style>body{font-family:Arial,sans-serif;padding:16px;color:#222}label{display:block;margin:12px 0 6px;font-weight:bold}textarea,input{width:100%;box-sizing:border-box;padding:9px;border:1px solid #aaa;border-radius:5px}textarea{min-height:95px}button{padding:12px;border:0;border-radius:6px;margin:12px 8px 0 0;cursor:pointer;font-weight:bold}.approve{background:#276842;color:white}.reject{background:#a72b32;color:white}.note{font-size:12px;color:#555}#message{margin-top:15px}</style></head><body><h3>Chapter content review</h3><label for="post">Post ID</label><input id="post" readonly><label for="name">Reviewer name</label><input id="name" placeholder="Your name" required><label for="comments">Comments (required for rejection)</label><textarea id="comments" placeholder="Feedback for the content team"></textarea><div><button class="approve" onclick="submitDecision('Approved')">Approve</button><button class="reject" onclick="submitDecision('Rejected')">Reject</button></div><div id="message" role="status"></div><p class="note">Submissions are saved to the RKL approval tracker.</p><script>document.getElementById('post').value=${JSON.stringify(post)};function submitDecision(decision){const postId=document.getElementById('post').value,name=document.getElementById('name').value.trim(),comments=document.getElementById('comments').value.trim(),msg=document.getElementById('message');if(!name){msg.textContent='Enter your name.';return}if(decision==='Rejected'&&!comments){msg.textContent='Please explain the rejection.';return}msg.textContent='Saving…';google.script.run.withSuccessHandler(()=>{msg.textContent=decision+' recorded successfully.'}).withFailureHandler(e=>{msg.textContent='Could not save: '+e.message}).saveRKLDecision({postId,name,comments,decision})}</script></body></html>`);
  return html.setTitle('RKL Content Review').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function saveRKLDecision(data) {
  const ss = SpreadsheetApp.openById(RKL_SHEET_ID);
  const posts = ss.getSheetByName('Posts');
  const responses = ss.getSheetByName('Approval Responses');
  const status = ss.getSheetByName('Approval Status');
  const postId = String(data.postId || '').trim();
  const decision = String(data.decision || '');
  const name = String(data.name || '').trim().slice(0,120);
  const comments = String(data.comments || '').trim().slice(0,3000);
  if (!/^RKL-\\d{3}$/.test(postId) || !['Approved','Rejected'].includes(decision) || !name || (decision === 'Rejected' && !comments)) throw new Error('Invalid review.');
  const rows = posts.getDataRange().getValues();
  const index = rows.findIndex((r,i) => i > 0 && r[0] === postId);
  if (index < 0) throw new Error('Unknown post ID.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const now = new Date();
    responses.appendRow([now, name, postId, decision, comments]);
    posts.getRange(index+1,8).setValue(decision);
    const sr = status.getDataRange().getValues();
    const si = sr.findIndex((r,i) => i > 0 && r[0] === postId);
    if (si >= 0) status.getRange(si+1,1,1,5).setValues([[postId,rows[index][2],decision,now,name]]);
    else status.appendRow([postId,rows[index][2],decision,now,name]);
  } finally {lock.releaseLock()}
  return {ok:true};
}
