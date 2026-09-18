import { b64DecodeUtf8, b64EncodeUtf8, clean } from './utils.js';

export class GitHubStore {
  constructor({owner,repo,branch='main',dataDir='data',token}){
    this.owner=clean(owner); this.repo=clean(repo); this.branch=clean(branch)||'main';
    this.dataDir=clean(dataDir).replace(/^\/+|\/+$/g,'')||'data'; this.token=clean(token);
    if(!this.owner||!this.repo) throw new Error('GitHub owner and repository are required.');
    if(!this.token) throw new Error('A fine-grained GitHub token is required for a private data repository.');
  }
  path(name){return `${this.dataDir}/${name}`.replace(/\/+/g,'/');}
  url(path){return `https://api.github.com/repos/${encodeURIComponent(this.owner)}/${encodeURIComponent(this.repo)}/contents/${path.split('/').map(encodeURIComponent).join('/')}`;}
  headers(){return {'Accept':'application/vnd.github+json','Authorization':`Bearer ${this.token}`,'X-GitHub-Api-Version':'2022-11-28','Content-Type':'application/json'};}
  async request(url,options={}){
    const res=await fetch(url,{...options,headers:{...this.headers(),...(options.headers||{})},cache:'no-store'});
    if(res.status===404) return {notFound:true,status:404};
    let body=null; try{body=await res.json();}catch{body={message:await res.text()};}
    if(!res.ok){const e=new Error(body?.message||`GitHub API error ${res.status}`);e.status=res.status;e.body=body;throw e;}
    return body;
  }
  async repoInfo(){return this.request(`https://api.github.com/repos/${encodeURIComponent(this.owner)}/${encodeURIComponent(this.repo)}`);}
  async getFile(name){
    const result=await this.request(`${this.url(this.path(name))}?ref=${encodeURIComponent(this.branch)}`);
    if(result?.notFound) return null;
    return {sha:result.sha,text:b64DecodeUtf8(result.content||''),path:result.path,htmlUrl:result.html_url};
  }
  async getJSON(name,fallback=null){const f=await this.getFile(name);if(!f)return fallback;try{return JSON.parse(f.text);}catch(e){throw new Error(`${name} contains invalid JSON: ${e.message}`);}}
  async putFile(name,text,{message=`Update ${name}`,sha=null}={}){
    const payload={message,content:b64EncodeUtf8(text),branch:this.branch}; if(sha)payload.sha=sha;
    return this.request(this.url(this.path(name)),{method:'PUT',body:JSON.stringify(payload)});
  }
  async putJSON(name,value,{message=`Update ${name}`,sha=null}={}){return this.putFile(name,JSON.stringify(value,null,2)+'\n',{message,sha});}
  async mutateJSON(name,fallback,mutator,message){
    for(let attempt=0;attempt<3;attempt++){
      const current=await this.getFile(name); let value=current?JSON.parse(current.text):structuredClone(fallback);
      const next=await mutator(structuredClone(value));
      try{return await this.putJSON(name,next,{message:message||`Update ${name}`,sha:current?.sha||null});}
      catch(e){if((e.status===409||e.status===422)&&attempt<2)continue;throw e;}
    }
    throw new Error(`Could not update ${name} because the file changed repeatedly. Refresh and try again.`);
  }
}
