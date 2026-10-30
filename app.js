/* Planta de Campo online: a casca (30/10/2026).
   A planta continua sendo o HTML gerado de cada obra, sem mudança. A casca:
   - faz o login Microsoft (app "Planta de Campo (protótipo)", Sites.Selected no site Montagem);
   - lê a lista de obras (obras.json no SharePoint) e abre a planta da obra num iframe;
   - guarda a planta para abrir sem sinal;
   - troca o compartilhar do Android (navigator.share) por gravação direta na Envios da obra.
     O "share" só termina depois que o SharePoint confirma o arquivo inteiro: a planta só marca
     "enviado" quando chegou de verdade. Sem sinal, espera o sinal voltar com a planta aberta.
   Os arquivos são os mesmos que o líder mandava pelo compartilhar: os importadores não mudam. */
var VERSAO = "0.1.5";
var CLIENTE = "be9bf767-d1a4-4735-adec-cf4aacc40391";
var LOCATARIO = "8999042a-e592-47e6-a75e-58f8901c3645";
var SITE_ID = "montaggiotech.sharepoint.com,b0667df5-71a6-4efd-a8b1-b02d9192ccb1,4d2f1e7c-b7be-4cf5-b3b0-12cb70c6febd";
var ESCOPOS = ["User.Read", "Sites.Selected"];
/* lista de obras: produção e teste (abrir com #teste liga o teste neste aparelho; #producao desliga) */
var MANIFESTOS = { producao: "General/_PLANTA ONLINE/obras.json", teste: "General/_TESTE PLANTA/obras.json" };
var PEDACO = 10 * 327680; /* a sessão de envio do Graph pede pedaços múltiplos de 320 KiB */

var $ = function(id){ return document.getElementById(id); };
function ler(k, padrao){ try{ var v = localStorage.getItem("online:" + k); return v == null ? padrao : JSON.parse(v); }catch(e){ return padrao; } }
function gravar(k, v){ try{ localStorage.setItem("online:" + k, JSON.stringify(v)); }catch(e){} }
function registro(t){ var l = ler("registro", []); l.push(new Date().toISOString() + " " + t); if(l.length > 80) l = l.slice(-80); gravar("registro", l); }
function hora(d){ return new Date(d).toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" }); }
function mb(b){ return (b / 1048576).toFixed(1).replace(".", ","); }
function espera(ms){ return new Promise(function(ok){ setTimeout(ok, ms); }); }
function tela(id){ ["t-entrar", "t-obras", "t-msg"].forEach(function(t){ $(t).hidden = t !== id; }); $("planta").hidden = $("faixa-obra").hidden = !!id; }
function mensagem(tit, txt, tentar){
  tela("t-msg"); $("msg-tit").textContent = tit; $("msg-txt").innerHTML = txt || "";
  $("msg-bt").hidden = !tentar; $("msg-bt").onclick = tentar || null;
}

(function(){ var h = location.hash;
  if(/teste/.test(h)) gravar("manifesto", "teste"); else if(/producao/.test(h)) gravar("manifesto", "producao"); })();
var MODO = ler("manifesto", "producao");
$("versao").textContent = $("versao2").textContent = "versão " + VERSAO + (MODO === "teste" ? " · MODO TESTE" : "");

/* ------------------------------------------------ login */
var BASE = location.origin + (location.pathname === "/" ? "" : location.pathname.replace(/index\.html$/, ""));
var pca = new msal.PublicClientApplication({
  auth: { clientId: CLIENTE, authority: "https://login.microsoftonline.com/" + LOCATARIO, redirectUri: BASE },
  cache: { cacheLocation: "localStorage" }
});
function conta(){ return pca.getAllAccounts()[0] || null; }
/* token sem tela; se a Microsoft pedir login de novo, o aviso mostra um botão (popup precisa de toque) */
function token(){
  return pca.acquireTokenSilent({ scopes: ESCOPOS, account: conta() }).then(function(r){ return r.accessToken; })
    .catch(function(e){
      if(!(e instanceof msal.InteractionRequiredAuthError)) throw e;
      return new Promise(function(ok, falha){
        aviso("A Microsoft pediu para entrar de novo antes de enviar.", "erro", [{ txt: "ENTRAR DE NOVO", g: true, fn: function(){
          pca.acquireTokenPopup({ scopes: ESCOPOS, account: conta() }).then(function(r){ aviso(null); ok(r.accessToken); }).catch(falha);
        } }]);
      });
    });
}

/* ------------------------------------------------ Graph */
function erroGraph(r){
  return r.json().catch(function(){ return {}; }).then(function(j){
    var e = new Error(r.status + " " + ((j.error && (j.error.code + ": " + j.error.message)) || r.statusText));
    e.status = r.status; e.codigo = j.error && j.error.code; throw e;
  });
}
function graph(caminho, opc){
  return token().then(function(tk){
    opc = opc || {}; opc.headers = opc.headers || {}; opc.headers.Authorization = "Bearer " + tk;
    return fetch("https://graph.microsoft.com/v1.0" + caminho, opc);
  }).then(function(r){ return r.ok ? (r.status === 204 ? null : r.json()) : erroGraph(r); });
}
function drive(){
  var d = ler("drive", null);
  if(d) return Promise.resolve(d);
  return graph("/sites/" + SITE_ID + "/drive?$select=id").then(function(j){ gravar("drive", j.id); return j.id; });
}
function rel(p){ return p.split("/").map(encodeURIComponent).join("/"); }
/* conteúdo de um arquivo pelo link de download do próprio Graph (já vem autenticado).
   O link só vem com o item inteiro: com $select o Graph não manda o @microsoft.graph.downloadUrl (30/10). */
function baixarTexto(p){
  var d0;
  return drive().then(function(d){ d0 = d; return graph("/drives/" + d + "/root:/" + rel(p)); })
    .then(function(it){
      var u = it["@microsoft.graph.downloadUrl"];
      var r = u ? fetch(u) : token().then(function(tk){ return fetch("https://graph.microsoft.com/v1.0/drives/" + d0 + "/items/" + it.id + "/content", { headers: { Authorization: "Bearer " + tk } }); });
      return r.then(function(x){ if(!x.ok) throw new Error("download " + x.status); return x.text(); })
        .then(function(t){ return { texto: t, eTag: it.eTag }; });
    });
}

/* ------------------------------------------------ obras */
function carregarObras(){
  return baixarTexto(MANIFESTOS[MODO]).then(function(r){ var m = JSON.parse(r.texto); gravar("obras:" + MODO, m); return m; })
    .catch(function(e){ var m = ler("obras:" + MODO, null); registro("lista de obras: " + e.message); if(m) return m; throw e; });
}
function mostrarObras(m){
  tela("t-obras");
  var c = conta(); $("quem").textContent = c ? "Conectado como " + c.name + " (" + c.username + ")" : "";
  $("lista-obras").innerHTML = "";
  (m.obras || []).forEach(function(o){
    var b = document.createElement("button"), atual = o.id === ler("ultima", null);
    b.className = "obra" + (atual ? " atual" : "");
    b.innerHTML = "<span></span><small></small>";
    b.firstChild.textContent = o.nome; b.lastChild.textContent = (o.lider ? "Líder: " + o.lider + " · " : "") + (o.periodo || "") + (atual ? " · aberta por último" : "");
    b.onclick = function(){ gravar("ultima", o.id); abrirObra(o); };
    $("lista-obras").appendChild(b);
  });
  if(!(m.obras || []).length) $("lista-obras").textContent = "Nenhuma obra liberada ainda.";
}

/* ------------------------------------------------ planta: guardada para abrir sem sinal */
function plantaGuardada(id){ return caches.open("plantas").then(function(c){ return c.match("planta/" + id); })
  .then(function(r){ return r ? r.text() : null; }); }
function guardarPlanta(id, html){ return caches.open("plantas").then(function(c){ return c.put("planta/" + id, new Response(html, { headers: { "Content-Type": "text/html" } })); }); }
function buscarPlanta(o){
  var meta = ler("planta:" + o.id, null);
  if(!navigator.onLine) return plantaGuardada(o.id);
  return drive().then(function(d){ return graph("/drives/" + d + "/root:/" + rel(o.planta) + "?$select=eTag"); }).then(function(it){
    if(meta && meta.eTag === it.eTag) return plantaGuardada(o.id).then(function(h){ return h || baixar(); });
    return baixar();
  }).catch(function(e){ registro("planta " + o.id + ": " + e.message); return plantaGuardada(o.id); });
  function baixar(){ return baixarTexto(o.planta).then(function(r){
    return guardarPlanta(o.id, r.texto).then(function(){ gravar("planta:" + o.id, { eTag: r.eTag, em: new Date().toISOString() }); return r.texto; }); }); }
}
/* o que entra na planta antes do código dela: o compartilhar vira gravação no SharePoint, e somem os passos do OneDrive */
var INJETAR = '<script>(function(){var P=parent.PlantaOnline;window.PLANTA_ONLINE=true;try{'
  + 'Object.defineProperty(navigator,"share",{configurable:true,value:function(d){return P.enviar((d&&d.files)||[]);}});'
  + 'Object.defineProperty(navigator,"canShare",{configurable:true,value:function(){return true;}});}catch(e){}'
  + 'document.addEventListener("DOMContentLoaded",function(){var n=document.querySelectorAll(".env-nota");'
  + 'for(var i=0;i<n.length;i++)n[i].textContent="Planta online: o arquivo vai direto para a pasta Envios da obra. O aviso verde no alto só aparece depois que o SharePoint confirmou.";'
  /* frases das plantas que mandam conferir no OneDrive: no modo online quem confere é o app */
  + 'var T=[[/ ?Confira no OneDrive → Envios se o arquivo está lá\\./g,""],[/ — confira na pasta Envios do OneDrive/g," — chegou no SharePoint"]];'
  + 'function lim(no){var w=document.createTreeWalker(no,4),t;while(t=w.nextNode()){var v=t.nodeValue,x=v;for(var j=0;j<T.length;j++)x=x.replace(T[j][0],T[j][1]);if(x!==v)t.nodeValue=x;}}'
  + 'lim(document.body);new MutationObserver(function(ms){for(var k=0;k<ms.length;k++)lim(ms[k].target);}).observe(document.body,{childList:true,subtree:true,characterData:true});});'
  + '})();<\/script><style>#env-passos .passo,#env-passos label.campo,#env-corpo ol{display:none!important}</style>';
var OBRA = null;
function abrirObra(o){
  OBRA = o;
  mensagem("Abrindo " + o.nome + "…", navigator.onLine ? "Baixando a planta da pasta da obra." : "Sem sinal: abrindo a última planta guardada.");
  buscarPlanta(o).then(function(html){
    if(!html){ mensagem("Não deu para abrir a planta", "Este aparelho ainda não tem a planta guardada e não conseguiu baixar. Precisa de sinal na primeira vez.", function(){ abrirObra(o); }); return; }
    var i = html.search(/<head[^>]*>/i), fim = i < 0 ? 0 : i + html.match(/<head[^>]*>/i)[0].length;
    tela(null); document.title = o.nome + " · Planta de Campo";
    $("faixa-nome").textContent = (MODO === "teste" ? "TESTE · " : "") + o.nome.replace(/^TESTE · /, "");
    $("planta").srcdoc = html.slice(0, fim) + INJETAR + html.slice(fim);
    registro("abriu " + o.id);
  });
}

/* ------------------------------------------------ envio: sessão de upload do Graph, em pedaços, retomável */
function aviso(txt, tipo, botoes, prog){
  var a = $("aviso");
  if(!txt){ a.hidden = true; return; }
  a.className = tipo || ""; a.hidden = false; a.innerHTML = "<div></div>";
  a.firstChild.innerHTML = txt;
  if(prog != null){ var b = document.createElement("div"); b.className = "barra"; b.innerHTML = "<i></i>"; b.firstChild.style.width = Math.round(prog * 100) + "%"; a.appendChild(b); }
  (botoes || []).forEach(function(x){ var bt = document.createElement("button"); bt.textContent = x.txt; if(x.g) bt.className = "g"; bt.onclick = x.fn; a.appendChild(bt); });
}
var avisoTimer = null;
function avisoSome(ms){ clearTimeout(avisoTimer); avisoTimer = setTimeout(function(){ aviso(null); }, ms); }
function sinal(){
  if(navigator.onLine) return Promise.resolve();
  aviso("<b>Sem sinal agora.</b> O envio sai sozinho quando o sinal voltar. Deixe a planta aberta. Nada foi perdido.", "erro");
  return new Promise(function(ok){ window.addEventListener("online", function f(){ window.removeEventListener("online", f); ok(); }); });
}
function enviarArquivo(f, destino, desistir){
  var total = f.size, url = null, tentativa = 0;
  function sessao(){
    return drive().then(function(d){
      return graph("/drives/" + d + "/root:/" + rel(destino) + ":/createUploadSession",
        { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "fail" } }) });
    }).then(function(j){ url = j.uploadUrl; });
  }
  /* já existe com o mesmo nome: se o tamanho bate, é o mesmo envio que chegou antes (toque repetido) */
  function jaExiste(){ return drive().then(function(d){ return graph("/drives/" + d + "/root:/" + rel(destino) + "?$select=size"); })
    .then(function(it){ if(it.size === total) return { size: it.size, repetido: true }; throw new Error("já existe um arquivo com esse nome e outro tamanho"); }); }
  function pedaco(ini){
    var fim = Math.min(ini + PEDACO, total);
    aviso("Enviando <b>" + f.name.replace(/</g, "&lt;") + "</b><br>" + mb(ini) + " de " + mb(total) + " MB", "", null, total ? ini / total : 0);
    return fetch(url, { method: "PUT", headers: { "Content-Range": "bytes " + ini + "-" + (fim - 1) + "/" + total }, body: f.slice(ini, fim) })
      .then(function(r){
        if(r.status === 200 || r.status === 201) return r.json();
        if(r.status === 202) return r.json().then(function(j){ tentativa = 0; return pedaco(proximo(j, fim)); });
        return erroGraph(r);
      }, function(e){ e.rede = true; throw e; })
      .catch(function(e){
        if(desistir.agora) throw e;
        if(e.rede || e.status >= 500 || e.status === 429){
          tentativa++;
          var s = Math.min(60, 2 * Math.pow(2, Math.min(tentativa, 5)));
          registro("envio " + f.name + ": " + (e.message || e) + " · nova tentativa em " + s + "s");
          aviso("A rede caiu no meio do envio. Tentando de novo em " + s + " s…", "erro",
            [{ txt: "SALVAR EM DOWNLOADS EM VEZ DISSO", fn: function(){ desistir.agora = true; } }]);
          return espera(s * 1000).then(sinal).then(retomar);
        }
        throw e;
      });
  }
  function proximo(j, padrao){ var n = j && j.nextExpectedRanges && j.nextExpectedRanges[0]; return n ? parseInt(n, 10) : padrao; }
  /* pergunta à sessão de onde continuar; se ela expirou, começa outra */
  function retomar(){
    if(desistir.agora) return Promise.reject(new Error("o líder preferiu salvar em Downloads"));
    return fetch(url).then(function(r){ return r.ok ? r.json() : null; }, function(){ return null; }).then(function(j){
      if(j) return pedaco(proximo(j, 0));
      return sessao().then(function(){ return pedaco(0); });
    });
  }
  return sessao().then(function(){ return pedaco(0); }, function(e){
    if(e.status === 409 || e.codigo === "nameAlreadyExists") return jaExiste();
    throw e;
  }).then(function(it){
    if(it.size !== total) throw new Error("chegou com tamanho diferente (" + it.size + " de " + total + ")");
    return it;
  });
}
var FILA = Promise.resolve();
window.PlantaOnline = {
  versao: VERSAO,
  enviar: function(arqs){
    var o = OBRA, desistir = { agora: false };
    ENVIANDO++;
    var p = FILA.then(function(){
      if(!o) throw new Error("obra não definida");
      return sinal().then(function(){
        return arqs.reduce(function(ac, f){
          return ac.then(function(l){ return enviarArquivo(f, o.envios + "/" + f.name, desistir).then(function(it){ l.push(it); return l; }); });
        }, Promise.resolve([]));
      });
    }).then(function(l){
      var b = 0; arqs.forEach(function(f){ b += f.size; });
      registro("enviado " + arqs.map(function(f){ return f.name; }).join(", ") + (l.some(function(x){ return x.repetido; }) ? " (já estava lá)" : ""));
      aviso("✓ <b>Chegou no SharePoint</b> às " + hora(new Date()) + "<br>" + arqs.length + " arquivo(s), " + mb(b) + " MB, na pasta Envios de " + o.nome.replace(/</g, "&lt;") + ".", "ok");
      avisoSome(9000);
    }, function(e){
      registro("falhou: " + (e.message || e));
      aviso("<b>Não chegou no SharePoint.</b> " + (e.message || e) + "<br>A planta vai salvar o arquivo em Downloads; nada foi perdido.", "erro",
        [{ txt: "OK", fn: function(){ aviso(null); } }]);
      var x = new Error(e.message || String(e)); x.name = "EnvioSharePoint"; throw x;
    });
    FILA = p.catch(function(){});
    FILA.then(function(){ ENVIANDO--; });
    return p;
  }
};

/* ------------------------------------------------ partida */
$("b-entrar").onclick = function(){ pca.loginRedirect({ scopes: ESCOPOS }); };
/* trocar de obra: a planta aberta já guardou tudo no aparelho; só não troca no meio de um envio */
var ENVIANDO = 0;
$("b-trocar").onclick = function(){
  if(ENVIANDO){ aviso("Espere o envio terminar para trocar de obra.", "erro"); avisoSome(4000); return; }
  carregarObras().then(mostrarObras).catch(function(){ var m = ler("obras:" + MODO, null); if(m) mostrarObras(m); });
};
$("b-sair").onclick = function(){ pca.logoutRedirect({ account: conta(), postLogoutRedirectUri: BASE }); };
function partida(){
  if(!conta()){ tela("t-entrar"); return; }
  mensagem("Carregando…", "");
  carregarObras().then(function(m){
    var u = ler("ultima", null), o = (m.obras || []).filter(function(x){ return x.id === u; })[0];
    if(o && !/obras/.test(location.hash)) abrirObra(o); else mostrarObras(m);
  }).catch(function(e){
    /* 404: a lista não existe (ainda não publicada); 403: a conta do líder não tem acesso ao site Montagem */
    var txt = e.status === 404 ? "A lista de obras ainda não foi publicada para este modo (" + MODO + "). Avise o Gustavo."
      : e.status === 403 ? "A sua conta ainda não tem acesso às pastas das obras no SharePoint. Avise o Gustavo: ele pede a liberação."
      : "Precisa de sinal na primeira vez.";
    mensagem("Não deu para ler a lista de obras", txt + "<br><small>Detalhe: " + String(e.message || e).replace(/</g, "&lt;") + "</small>", partida);
  });
}
pca.initialize().then(function(){ return pca.handleRedirectPromise(); }).then(partida).catch(function(e){
  registro("login: " + (e.errorCode || "") + " " + (e.message || e));
  tela("t-entrar"); aviso("O login não terminou: " + (e.errorMessage || e.message || e), "erro", [{ txt: "OK", fn: function(){ aviso(null); } }]);
});
if("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(function(e){ registro("sw: " + e.message); });
