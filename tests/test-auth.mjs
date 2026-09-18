import { createHash } from 'node:crypto';
import { verifyPassword, hashPassword } from '../js/auth.js';
const password='ExamplePass123!'; const salt='ABC123';
const legacy=createHash('sha256').update(`${salt}|${password}`,'utf8').digest('base64url');
if(!(await verifyPassword(password,{passwordHash:legacy,passwordSalt:salt})))throw new Error('Legacy hash verification failed');
const h=await hashPassword(password);
if(!(await verifyPassword(password,{passwordHash:h.hash,passwordSalt:h.salt,passwordIterations:h.iterations,passwordAlgorithm:h.algorithm})))throw new Error('PBKDF2 verification failed');
console.log('auth ok');
