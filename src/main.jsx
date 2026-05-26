import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'

async function initStorage() {
  const SUPABASE_URL      = import.meta.env.VITE_SUPABASE_URL;
  const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;
  if (SUPABASE_URL && SUPABASE_ANON_KEY) {
    const h = {'Content-Type':'application/json','apikey':SUPABASE_ANON_KEY,'Authorization':`Bearer ${SUPABASE_ANON_KEY}`};
    const base = `${SUPABASE_URL}/rest/v1/cf_storage`;
    window.storage = {
      get: async k => { const r=await fetch(`${base}?key=eq.${encodeURIComponent(k)}&select=key,value`,{headers:h}); const d=await r.json(); return d?.[0]?{key:k,value:d[0].value}:null; },
      set: async (k,v) => { await fetch(base,{method:'POST',headers:{...h,'Prefer':'resolution=merge-duplicates'},body:JSON.stringify({key:k,value:v,updated_at:new Date().toISOString()})}); return {key:k,value:v}; },
      delete: async k => { await fetch(`${base}?key=eq.${encodeURIComponent(k)}`,{method:'DELETE',headers:h}); return {key:k,deleted:true}; },
      list: async (p='') => { const r=await fetch(p?`${base}?key=like.${encodeURIComponent(p+'%')}&select=key`:`${base}?select=key`,{headers:h}); const d=await r.json(); return {keys:Array.isArray(d)?d.map(x=>x.key):[]}; },
    };
  } else {
    window.storage = {
      get: async k => { const v=localStorage.getItem(k); return v!=null?{key:k,value:v}:null; },
      set: async (k,v) => { try{localStorage.setItem(k,v);}catch{} return {key:k,value:v}; },
      delete: async k => { localStorage.removeItem(k); return {key:k,deleted:true}; },
      list: async (p='') => { const keys=[]; for(let i=0;i<localStorage.length;i++){const k=localStorage.key(i);if(k&&k.startsWith(p))keys.push(k);} return {keys}; },
    };
  }
}

initStorage().then(()=>{
  ReactDOM.createRoot(document.getElementById('root')).render(
    <React.StrictMode><App/></React.StrictMode>
  );
});
