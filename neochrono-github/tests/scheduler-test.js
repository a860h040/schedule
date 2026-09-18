import { emptyDatabase } from '../assets/js/seed.js';
import { validateConfiguration } from '../assets/js/rules.js';
const db=emptyDatabase();
const v=validateConfiguration(db);
if(!v.errors.some(x=>x.includes('7-on/7-off'))) throw new Error('Expected missing 7-on/7-off configuration to be detected.');
console.log('NeoChrono core validation test passed.');
