// All Apps Script traffic is intercepted before navigation. No live sheet requests.
const fs=require('fs'),vm=require('vm'),assert=require('node:assert/strict');
const {pathToFileURL}=require('url');
const {chromium}=require('/Users/bhonepyaesone/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');
const source=fs.readFileSync('tests/schema.test.cjs','utf8');
const setup=source.slice(source.indexOf('  const tables ='),source.indexOf("  assert.equal(post('addDebt'"));
const sandbox={require,fs,vm,legacy:false};vm.createContext(sandbox);
vm.runInContext(setup+';globalThis.api=context;',sandbox);
(async()=>{
 const browser=await chromium.launch({headless:true,channel:'chrome'});
 try {
 const page=await browser.newPage({viewport:{width:390,height:844}});
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 const requests=[];
 await page.route(/https:\/\/script\.(google|googleusercontent)\.com\//,async route=>{
 const req=route.request();let result;
 if(req.method()==='POST') {const body=req.postData();requests.push(JSON.parse(body).action);result=sandbox.api.doPost({postData:{contents:body}});}
 else result=sandbox.api.doGet();
 await route.fulfill({json:result,headers:{'access-control-allow-origin':'*'}});
 });
 await page.goto(pathToFileURL(require('path').resolve('index.html')).href);
 await page.waitForFunction(()=>document.getElementById('toastMsg').textContent==='Synced!');
 async function saved(message){await page.waitForFunction(m=>document.getElementById('toastMsg').textContent===m,message);}
 await page.locator('header button[onclick="openAddModal()"]').click();
 await page.locator('#addName').fill('TEST Student');await page.locator('#addYear').selectOption('Year 9A');await page.locator('#addAmount').fill('1000');
 await page.locator('#addDebtForm button[type=submit]').click();await saved('Bill saved!');
 const debt=()=>sandbox.api.doGet().debts[0];assert.equal(debt().balance,1000);
 await page.locator('#debtsContainer button').filter({hasText:'PAY'}).click();await page.locator('#payAmount').fill('300');await page.locator('#paymentModal button[type=submit]').click();await saved('Payment confirmed!');assert.equal(debt().balance,700);
 await page.locator('#debtsContainer button').filter({hasText:'Add'}).click();await page.locator('#incAmount').fill('200');await page.locator('#increaseModal button[type=submit]').click();await saved('Bill increased!');assert.equal(debt().balance,900);
 assert.equal(await page.locator('#dashTodayUnpaid').innerText(),'K 1,200');
 await page.locator('#debtsContainer button').filter({hasText:'History'}).click();await page.locator('#editStudentDetails').click();await page.locator('[name=studentName]').fill('TEST Corrected');await page.locator('#profileEditor [name=year]').fill('Year 10A');await page.locator('#profileEditor [name=category]').selectOption('Kitchen');await page.locator('#profileEditor button[type=submit]').click();await saved('Student details updated');assert.equal(debt().name,'TEST Corrected');
 const payRow=page.locator('#detailHistoryContainer > div').filter({hasText:'Payment'});await payRow.getByRole('button').click();await page.locator('#transactionEditor [name=amount]').fill('400');await page.locator('#transactionEditor button[type=submit]').click();await saved('Transaction updated');assert.equal(debt().balance,800);
 await page.locator('#debtsContainer button').filter({hasText:'PAY'}).click();await page.locator('#payAmount').fill('800');await page.locator('#paymentModal button[type=submit]').click();await saved('Payment confirmed!');assert.equal(debt().status,'Paid');
 await page.locator('#tab-paid').click();assert.match(await page.locator('#debtsContainer').innerText(),/TEST Corrected/);
 await page.locator('#debtsContainer button').filter({hasText:'History'}).click();
 page.on('dialog',d=>d.accept());
 await page.locator('#detailHistoryContainer > div').filter({hasText:'K 800'}).getByRole('button').click();await page.locator('#transactionEditor [data-action=delete]').click();await saved('Transaction deleted');assert.equal(debt().balance,800);
 await page.locator('#tab-unpaid').click();await page.locator('#searchInput').fill('nonexistent');await page.locator('#searchInput').press('x');assert.equal(await page.locator('#debtsContainer > div').count(),0);
 await page.locator('#searchInput').fill('');await page.locator('#searchInput').press('Backspace');
 await page.locator('#typeFilter').selectOption('Canteen');assert.equal(await page.locator('#debtsContainer > div').count(),0);await page.locator('#typeFilter').selectOption('Kitchen');assert.equal(await page.locator('#debtsContainer > div').count(),1);
 await page.locator('header button').filter({hasText:'HISTORY'}).click();await page.waitForTimeout(400);assert.match(await page.locator('#globalHistoryList').innerText(),/TEST Corrected/);
 assert.deepEqual(errors,[]);console.log('PASS browser: create, partial pay, increase, dashboard, student edit, payment edit, full pay, paid tab, delete payment, search, category filter, global history.');console.log('Mock actions:',requests.join(', '));
 } finally {await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
