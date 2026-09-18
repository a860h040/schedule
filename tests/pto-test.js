import { reconcilePtoAutoApprovals } from '../assets/js/rules.js';
const base=[1,2,3].map((n,i)=>({'Record Type':'PTO','Record ID':'P'+n,Pharmacist:'A'+n,Username:'U'+n,'Start Date':'2026-10-01','End Date':'2026-10-01',Status:'Pending','Submitted At':`2026-09-0${i+1}T10:00:00Z`,'Reviewed By':''}));
let r=reconcilePtoAutoApprovals(base,2);
if(r[0].Status!=='Approved'||r[1].Status!=='Approved'||r[2].Status!=='Pending')throw new Error('First-two PTO rule failed.');
r[0].Status='Rejected';r=reconcilePtoAutoApprovals(r,2);
if(r[2].Status!=='Approved')throw new Error('PTO promotion after rejection failed.');
console.log('PTO auto-approval test passed.');
