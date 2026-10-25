// Botão "Enviar boletos ao painel" (Sistema Insanos > tabela de mensalidades).
// Este arquivo é carregado pelo favorito, que traz o endereço e a chave da planilha (não guarde chaves aqui).
(function(){
  var C=window.__INSANOS_BOLETOS__||{};
  if(!C.url||!C.chave){alert("Favorito incompleto: copie de novo o botão no painel (Pagamentos).");return}
  (function(URL_,CHAVE_){
    const MESES=["janeiro","fevereiro","marco","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];
    const nz=s=>String(s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/\s+/g," ").trim();
    const txt=c=>c?(c.innerText||c.textContent||""):"";
    const grade=r=>{const g=[];let c=0;for(const cel of r.cells){const n=cel.colSpan||1;for(let k=0;k<n;k++)g[c+k]=cel;c+=n}return g};
    const docs=[];const junta=w=>{try{docs.push(w.document);for(let k=0;k<w.frames.length;k++)junta(w.frames[k])}catch(_){}};junta(window);
    const ehNome=r=>/^\s*\d+\s*[-–]\s*\S/.test(txt(r.cells[0]));
    let hdr=null,tabs=[];
    for(const d of docs){const ts=[...d.querySelectorAll("table")];tabs=tabs.concat(ts);
      for(const tb of ts){for(const r of tb.rows){const tx=[...r.cells].map(c=>nz(txt(c)));if(tx.some(x=>x.startsWith("janeiro"))&&tx.some(x=>x.startsWith("dezembro"))){hdr=r;break}}if(hdr)break}if(hdr)break}
    const linhas=tabs.flatMap(tb=>[...tb.rows]).filter(r=>r.cells.length>=13&&ehNome(r));
    const folhas=[];for(const d of docs){for(const el of d.querySelectorAll("body *")){if(el.children.length)continue;const s=txt(el).trim();if(s&&s.length<40)folhas.push(el)}}
    const caixa=el=>{const r=el.getBoundingClientRect();return {x:(r.left+r.right)/2,y:(r.top+r.bottom)/2,l:r.left,r:r.right,h:r.height}};
    let porPos=null;
    // os textos soltos da página (mesmo dentro de blocos com ícones ou outros elementos junto), com a posição exata de cada um
    const textos=[];for(const d of docs){const tw=d.createTreeWalker(d.body,NodeFilter.SHOW_TEXT);let n;while((n=tw.nextNode())){const s=(n.nodeValue||"").trim();if(!s||s.length>40)continue;
      const rg=d.createRange();rg.selectNodeContents(n);const r=rg.getBoundingClientRect();if(!r.width&&!r.height)continue;textos.push({s,el:n.parentElement,b:{x:(r.left+r.right)/2,y:(r.top+r.bottom)/2,l:r.left,r:r.right,h:r.height}})}}
    if(!linhas.length){
      let cabs=MESES.map(m=>{const c=textos.find(o=>nz(o.s)===m||nz(o.s).startsWith(m+" ")||nz(o.s).startsWith(m.slice(0,3)+"/"));return c?c.b:null});
      const nomes=textos.filter(o=>/^\d+\s*[-–]\s*\S.*$/.test(o.s)&&!/\d+,\d{2}/.test(o.s)).map(o=>({el:o.el,s:o.s,b:o.b}));
      if(cabs.filter(Boolean).length<10&&nomes.length){
        // plano B: as 12 colunas pela posição dos próprios links (1º grupo da esquerda = janeiro …)
        const xs=docs.flatMap(d=>[...d.querySelectorAll("a[href]")]).filter(a=>/^https?:/i.test(a.href)&&txt(a).trim()).map(a=>caixa(a)).filter(b=>b.x>Math.max(...nomes.map(n=>n.b.r))).map(b=>b.x).sort((a,b)=>a-b);
        const grupos=[];for(const x of xs){const g=grupos[grupos.length-1];if(g&&x-g.f<25){g.f=x;g.s+=x;g.n++}else grupos.push({i:x,f:x,s:x,n:1})}
        if(grupos.length>=10&&grupos.length<=12){cabs=grupos.map(g=>({x:g.s/g.n,l:g.i-20,r:g.f+20}));while(cabs.length<12)cabs.push(null)}}
      if(cabs.filter(Boolean).length>=10&&nomes.length)porPos={cabs,nomes,plano:cabs.length&&cabs.every(Boolean)?"":"B"}}
    if(!linhas.length&&!porPos){const ex=textos.filter(o=>/jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez/i.test(o.s)).slice(0,4).map(o=>o.s).join(" | ");
      alert("Não achei a tabela de mensalidades nesta página.\n(diagnóstico: "+tabs.length+" tabela(s), "+(docs.length-1)+" quadro(s), "+textos.filter(o=>MESES.includes(nz(o.s))).length+" mês(es) e "+textos.filter(o=>/^\d+\s*[-–]\s*\S/.test(o.s)).length+" nome(s) achados; textos com mês: "+(ex||"nenhum")+")");return}
    const hoje=new Date(),resp=prompt("Pegar os boletos de qual mês? (1 a 12)",String(hoje.getMonth()+1));if(resp===null)return;
    const mi=parseInt(resp,10)-1;if(!(mi>=0&&mi<12)){alert("Mês inválido.");return}
    if(porPos){
      const cabs=porPos.cabs,meio=k=>cabs[k]?cabs[k].x:null;
      const colDe=x=>{let best=-1,dist=1e9;for(let k=0;k<12;k++){const c=cabs[k];if(!c)continue;if(x>=c.l-2&&x<=c.r+2)return k;const dd=Math.abs(x-c.x);if(dd<dist){dist=dd;best=k}}return best};
      const altura=Math.max(12,...porPos.nomes.map(n=>n.b.h));
      const links=docs.flatMap(d=>[...d.querySelectorAll("a[href]")]).filter(a=>/^https?:/i.test(a.href)&&txt(a).trim());
      const itens=[];const vistos=new Set();
      for(const a of links){const b=caixa(a);if(colDe(b.x)!==mi)continue;
        let dono=null,dy=1e9;for(const n of porPos.nomes){const d=Math.abs(n.b.y-b.y);if(d<dy){dy=d;dono=n}}
        if(!dono||dy>altura*0.8)continue;const m=String(dono.s||txt(dono.el)).trim().match(/^\s*\d+\s*[-–]\s*(.+?)\s*$/);if(!m||vistos.has(m[1]))continue;
        vistos.add(m[1]);itens.push({ap:m[1],link:a.href,status:txt(a).trim(),valor:""})}
      const semLink=porPos.nomes.length-itens.length;
      return enviar(itens,semLink,"leitura pela posição na tela")}
    let ini=-1,fim=-1;
    if(hdr){let c=0;for(const cel of hdr.cells){const n=cel.colSpan||1;if(nz(txt(cel)).startsWith(MESES[mi])){ini=c;fim=c+n-1}c+=n}}
    if(ini<0){const tot=grade(linhas[0]).length,por=Math.max(1,Math.floor((tot-1)/12));ini=1+mi*por;fim=ini+por-1}
    const itens=[];let semLink=0;
    function enviar(itens,semLink,como){
      if(!itens.length){alert("Nenhum boleto com link no mês "+(mi+1)+".\n(diagnóstico: "+como+")");return}
      const ano=hoje.getFullYear()+(mi<hoje.getMonth()-6?1:mi>hoje.getMonth()+6?-1:0);
      const mes=ano+"-"+String(mi+1).padStart(2,"0");
      if(!confirm("Enviar ao painel "+itens.length+" boleto(s) de "+MESES[mi]+"?"+(semLink>0?"\n("+semLink+" irmão(s) sem link nesse mês)":"")+"\n\nExemplo: "+itens[0].ap+" - "+itens[0].status))return;
      const pac={chave:CHAVE_,quem:"botão do Sistema Insanos",dados:{tipo:"boletos_sistema",mes:mes,itens:itens}};
      const porFormulario=()=>{
        const fm=document.createElement("form");fm.method="POST";fm.action=URL_;fm.target="_blank";fm.enctype="text/plain";fm.acceptCharset="UTF-8";fm.style.display="none";
        const ip=document.createElement("input");ip.type="hidden";
        const js=JSON.stringify(pac);ip.name=js.slice(0,-1)+',"_":"';ip.value='"}';
        fm.appendChild(ip);document.body.appendChild(fm);fm.submit();setTimeout(()=>fm.remove(),3000);
        alert("✅ Enviado! Abriu uma aba com {\"ok\":true}: pode fechar. Em uns 3 minutos os links estão na planilha.")};
      fetch(URL_,{method:"POST",mode:"no-cors",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(pac)})
        .then(()=>alert("✅ Enviado! Em uns 3 minutos os links estão na planilha e o painel atualiza."))
        .catch(()=>{try{porFormulario()}catch(e){alert("Não consegui enviar ("+e.message+"). Avise o Edu.")}})}
    for(const r of linhas){const g=grade(r),m=txt(g[0]).trim().match(/^\s*\d+\s*[-–]\s*(.+?)\s*$/);if(!m)continue;
      let a=null,st="",val="";for(let k=ini;k<=fim;k++){const cel=g[k];if(!cel)continue;const x=cel.querySelector("a[href]");if(x&&!a){a=x}const tt=txt(cel).trim();if(/\d+,\d{2}/.test(tt)&&!val)val=tt;else if(tt&&!st)st=tt}
      if(a&&/^https?:/i.test(a.href))itens.push({ap:m[1],link:a.href,status:(txt(a)||st).trim(),valor:val});else semLink++}
    enviar(itens,semLink,linhas.length+" irmão(s), colunas "+ini+" a "+fim+(hdr?", cabeçalho achado":", sem cabeçalho"))})(C.url,C.chave);
})();
