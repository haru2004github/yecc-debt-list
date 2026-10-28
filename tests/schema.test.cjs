const fs = require('fs');
const vm = require('vm');
const assert = require('node:assert/strict');
const html = fs.readFileSync('index.html', 'utf8');
const scripts = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)].map(m => m[1]);
scripts.forEach(source => new vm.Script(source));
const frontend = { window: {} };
vm.createContext(frontend);
vm.runInContext(scripts.at(-1), frontend);
vm.runInContext(`normalizeSheetData({debts:[{debt_id:'D-1',id:'MANUAL-unstable',name:'Student',total_debt:800,paid_amount:100,balance:700}],payments:[{pyament_id:'P-1',payment_date:'2026-10-01T00:00:00Z',amount_paid:100,debt_id:'D-1'}]});`, frontend);
assert.equal(vm.runInContext('debtsData[0].id',frontend),'D-1');
assert.equal(vm.runInContext('paymentsData[0].amount',frontend),100);
assert.equal(vm.runInContext('paymentsData[0].date',frontend),'2026-10-01T00:00:00Z');

for (const legacy of [false,true]) {
  const tables = {
    Debts: [legacy ? ['id','user_id','date','name','year','total_debt','paid_amount','balance','status','category'] : ['debt_id','date','name','year','total_debt','paid_amount','balance','status','category']],
    Payments: [legacy ? ['id','debt_id','date','amount'] : ['pyament_id','debt_id','payment_date','amount_paid']]
  };
  const sheets = Object.fromEntries(Object.entries(tables).map(([name,rows]) => [name,{
    getLastRow:()=>rows.length,
    getDataRange:()=>({getValues:()=>rows.map(row=>row.slice())}),
    appendRow: row=>rows.push(row),
    deleteRow: r=>rows.splice(r-1,1),
    getRange:(r,c)=>({setValue:v=>{rows[r-1][c-1]=v;}})
  }]));
  const context = {
    SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:n=>sheets[n]})},
    ContentService:{MimeType:{JSON:'json'},createTextOutput:text=>({setMimeType:()=>JSON.parse(text)})},
    LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>{}})},
    Utilities:{getUuid:()=> 'uuid'}
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync('Code.gs','utf8'),context);
  const post=(action,data)=>context.doPost({postData:{contents:JSON.stringify({action,data})}});
  assert.equal(post('addDebt',{id:'D-1',name:'Student',total_debt:800}).status,'success');
  assert.equal(post('addPayment',{id:'P-1',debt_id:'D-1',amount:100}).status,'success');
  assert.equal(post('increaseDebt',{id:'INC-1',debt_id:'D-1',amount:200}).status,'success');
  let data=context.doGet();
  assert.equal(data.debts[0].id,'D-1');
  assert.equal(data.debts[0].total_debt,1000);
  assert.equal(data.debts[0].paid_amount,100);
  assert.equal(data.debts[0].balance,900);
  assert.equal(data.payments[1].amount,-200);
  const before=JSON.stringify(tables);
  assert.equal(post('addPayment',{id:'P-2',debt_id:'missing',amount:100}).status,'error');
  assert.equal(post('addPayment',{id:'P-1',debt_id:'D-1',amount:100}).status,'error');
  assert.equal(post('addPayment',{id:'P-3',debt_id:'D-1',amount:10000}).status,'error');
  assert.equal(JSON.stringify(tables),before);
  assert.equal(post('addPayment',{id:'P-4',debt_id:'D-1',amount:900}).status,'success');
  assert.equal(context.doGet().debts[0].status,'Paid');
  const edit = (kind,id,amount,deleting=false) => {
    const current = context.doGet();
    const record = (kind === 'initial' ? current.debts : current.payments).find(r=>r.id===id);
    const expected = kind === 'initial' ? record.total_debt-current.payments.filter(p=>p.debt_id===id && p.amount<0).reduce((s,p)=>s-p.amount,0) : record.amount;
    return post(deleting?'deleteTransaction':'updateTransaction',{kind,id,amount,date:'2026-10-08T12:00:00Z',expected_amount:expected,expected_date:new Date(record.date).toISOString()});
  };
  assert.equal(edit('payment','P-4',800).status,'success');
  assert.equal(context.doGet().debts[0].balance,100);
  assert.equal(edit('payment','INC-1',100).status,'success');
  assert.equal(context.doGet().debts[0].balance,0);
  assert.equal(edit('payment','INC-1',0,true).status,'error');
  assert.equal(edit('initial','D-1',0,true).status,'error');
  assert.equal(edit('payment','P-4',0,true).status,'success');
  assert.equal(context.doGet().debts[0].balance,800);
  assert.equal(edit('payment','INC-1',0,true).status,'success');
  assert.equal(edit('initial','D-1',1000).status,'success');
  assert.equal(context.doGet().debts[0].balance,900);
  assert.equal(post('updateTransaction',{kind:'initial',id:'D-1',amount:5,expected_amount:800,expected_date:'2026-10-08T12:00:00Z'}).status,'error');
  assert.equal(edit('payment','P-1',0,true).status,'success');
  const beforeProfile = context.doGet().debts[0];
  const profile = {id:'D-1', name:'Corrected Student',year:'Year 10A',category:'Kitchen',expected:{name:beforeProfile.name,year:beforeProfile.year,category:beforeProfile.category}};
  assert.equal(post('updateDebtDetails',profile).status,'success');
  const afterProfile = context.doGet().debts[0];
  assert.equal(afterProfile.name,'Corrected Student');
  assert.equal(afterProfile.year,'Year 10A');
  assert.equal(afterProfile.category,'Kitchen');
  for(const key of ['id','total_debt','paid_amount','balance']) assert.equal(afterProfile[key],beforeProfile[key]);
  assert.equal(post('updateDebtDetails',profile).status,'error');
  assert.equal(edit('initial','D-1',0,true).status,'success');
  assert.equal(context.doGet().debts.length,0);

}
console.log('PASS: frontend syntax and aliases; both sheet schemas; add, pay, increase, validation and paid status.');

// Existing borrowers, multiple increases, same-day creation, later increases,
// historical dates and category filters must all use transaction dates.
const elements = {};
frontend.document = { getElementById: id => elements[id] ||= {} };
elements.dashDate = { value: '2026-10-06' };
vm.runInContext(`normalizeSheetData({
  debts: [
    {id:'old',date:'2026-10-01T10:00:00',total_debt:1500,balance:1300,category:'Canteen'},
    {id:'new',date:'2026-10-06T10:00:00',total_debt:1000,balance:1000,category:'Canteen'},
    {id:'kitchen',date:'2026-10-01T10:00:00',total_debt:500,balance:500,category:'Kitchen'}
  ],
  payments: [
    {debt_id:'old',date:'2026-10-06T11:00:00',amount:-200},
    {debt_id:'old',date:'2026-10-06T12:00:00',amount:-300},
    {debt_id:'old',date:'2026-10-06T13:00:00',amount:200},
    {debt_id:'new',date:'2026-10-06T14:00:00',amount:-100},
    {debt_id:'new',date:'2026-10-07T14:00:00',amount:-300},
    {debt_id:'kitchen',date:'2026-10-06T14:00:00',amount:-100}
  ]
}); updateDashboard();`, frontend);
assert.equal(elements.dashTodayUnpaid.innerText,'K 1,300');
assert.equal(elements.dashTodayBorrowers.innerText,3);
assert.equal(elements.dashTodayCollected.innerText,'K 200');
vm.runInContext("currentDashboardView = 'Canteen'; updateDashboard();",frontend);
assert.equal(elements.dashTodayUnpaid.innerText,'K 1,200');
assert.equal(elements.dashTodayBorrowers.innerText,2);
elements.dashDate.value = '2026-10-01';
vm.runInContext('updateDashboard()',frontend);
assert.equal(elements.dashTodayUnpaid.innerText,'K 1,000');
elements.dashDate.value = '2026-10-07';
vm.runInContext('updateDashboard()',frontend);
assert.equal(elements.dashTodayUnpaid.innerText,'K 300');
assert.equal(elements.dashTodayBorrowers.innerText,1);
console.log('PASS: daily borrowing amounts, unique debt counts, historical dates and category filters.');
