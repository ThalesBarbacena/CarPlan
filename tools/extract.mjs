import fs from 'fs';
const html = fs.readFileSync(process.argv[2],'utf8');
let js = html.match(/<script type="module">([\s\S]*?)<\/script>/)[1];
js = js.replace(/^import .*$/mg,'');
const pre = `
const handler={ get(t,k){ if(k===Symbol.toPrimitive) return ()=>10; if(k===Symbol.iterator) return function*(){}; if(k==='then') return undefined; return P; },
  set(){return true}, apply(){return P}, construct(){return P}, has(){return true} };
const P=new Proxy(function(){},handler);
const THREE=P, OrbitControls=P, CSS2DRenderer=P, CSS2DObject=P;
const document=P, getComputedStyle=()=>P, matchMedia=()=>P, MutationObserver=P, ResizeObserver=P, localStorage=P, devicePixelRatio=1;
const requestAnimationFrame=()=>{}; const window=globalThis; globalThis.addEventListener=()=>{}; const location={hash:''};
`;
const fn = new Function(pre + js + '\nreturn window.__ITEMS;');
const items = fn();
fs.writeFileSync(process.argv[3], JSON.stringify(items,null,1));
console.log('items', items.length, 'activities', items.reduce((s,i)=>s+i.stages.length,0));
const byCode={}; items.forEach(i=>byCode[i.code]=(byCode[i.code]||0)+1); console.log(byCode);
console.log(items.slice(0,5).map(i=>i.id).join(' '));
