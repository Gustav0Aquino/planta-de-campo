/* Planta de Campo online, casca 0.3: o NÚCLEO (18/11/2026, Claude).
   Desenho: ../DESENHO_CASCA_0.3.md (seções 0 e 10).

   Divisão do trabalho:
   - este arquivo decide (dados, quem é quem, estado do envio, espaço) e devolve DADOS, nunca HTML;
   - as telas, os textos e o fluxo são do design (Codex), que chama window.Nucleo.*;
   - app.js só liga os ganchos (enviar, abrir obra) e mostra uma faixa provisória.

   Regras que nenhuma tela muda:
   - verde só com confirmação do SharePoint de TODAS as partes e nada pendente na planta;
   - vermelho (envio que começou e não terminou) não some sozinho;
   - sem registro de editores, a obra abre como hoje (editor). O modo leitura só vale
     quando o registro existe e a conta não é a editora;
   - identidade = conta Microsoft (e-mail). Nome serve só para histórico "não confirmado".

   Permissões no SharePoint ficaram EM ABERTO (18/11): este núcleo não muda permissão de nada. */
(function(){
"use strict";

/* ------------------------------------------------ utilidades */
function norm(t){ return String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/\s+/g, " ").trim(); }
function mesmoEmail(a, b){ return !!a && !!b && norm(a) === norm(b); }
/* nome do líder ("Fábio") contra o nome da conta ("Fabio Souza"): primeiro nome igual. Só para "não confirmado". */
function mesmoNome(lider, nomeConta){
  var a = norm(lider).split(" ")[0], b = norm(nomeConta).split(" ")[0];
  return !!a && a.length > 1 && a === b;
}
/* sem lista de acessos (modo da 0.2): o nome escrito pode ser o primeiro nome OU um sobrenome da conta */
function algumNome(lider, nomeConta){
  var a = norm(lider); if(!a || a.length < 2) return false;
  return mesmoNome(lider, nomeConta) || norm(nomeConta).split(" ").indexOf(a) >= 0;
}
function hojeISO(){ return new Date().toISOString().slice(0, 10); }
function lerLS(k, padrao){ try{ var v = localStorage.getItem("online:" + k); return v == null ? padrao : JSON.parse(v); }catch(e){ return padrao; } }
function gravarLS(k, v){ try{ localStorage.setItem("online:" + k, JSON.stringify(v)); }catch(e){} }
/* pasta do manifesto (produção ou teste): acessos.json e editores.json moram ao lado do obras.json */
function aoLado(manifestoCaminho, nome){ return manifestoCaminho.replace(/[^/]+$/, nome); }

/* ------------------------------------------------ acessos, registro de editores e trocas de líder
   acessos.json : {"versao":1, "adms":[{nome,email}], "lideres":[{nome,email}]}
                  Sem ele, o app se comporta como a 0.2 (todo mundo marca). Com ele (Gustavo, 18/11):
                  adm vê todas as obras ativas em consulta; líder marca a obra dele e consulta as outras;
                  conta fora da lista só consulta.
   editores.json: {"versao":1, "obras":{"<id>":{"editor":{nome,email,desde,autorizado_por},
                                                 "historico":[{nome,email,de,ate,papel,autorizado_por}]}}}  (opcional)
   trocas/       : um arquivo por "Agora sou o líder", nunca sobrescrito:
                   <obraId>__<AAAAMMDD-HHMMSS>.json  {obra, nome, email, em, anterior}
                   Quem trocou é o "Criado por" do SharePoint (não dá para forjar pelo nome do arquivo).
                   A troca mais recente de cada obra vale mais que editores.json e que o líder do obras.json. */
function lerOpcional(baixar, caminho, chaveCache){
  return baixar(caminho).then(function(r){
    var j = JSON.parse(String(r.texto).replace(/^﻿/, ""));
    gravarLS(chaveCache, { dados: j, eTag: r.eTag, em: new Date().toISOString() });
    return { dados: j, eTag: r.eTag, origem: "sharepoint" };
  }).catch(function(e){
    var c = lerLS(chaveCache, null);
    if(e && e.status === 404) return { dados: null, eTag: null, origem: "nao_existe" };
    if(c) return { dados: c.dados, eTag: c.eTag, origem: "cache", em: c.em };
    return { dados: null, eTag: null, origem: "erro", erro: String(e && e.message || e) };
  });
}
function carregarAcessos(baixar, manifestoCaminho){ return lerOpcional(baixar, aoLado(manifestoCaminho, "acessos.json"), "acessos:" + manifestoCaminho); }
function carregarEditores(baixar, manifestoCaminho){ return lerOpcional(baixar, aoLado(manifestoCaminho, "editores.json"), "editores:" + manifestoCaminho); }
function pastaTrocas(manifestoCaminho){ return aoLado(manifestoCaminho, "trocas"); }
/* listar(pasta) -> Promise<[{name, createdDateTime, createdBy}]>; pasta que não existe = nenhuma troca */
function carregarTrocas(listar, manifestoCaminho){
  return listar(pastaTrocas(manifestoCaminho)).then(function(itens){
    var t = (itens || []).map(function(x){
      var m = /^(.+)__(\d{8}-\d{6})(?:-\d+)?\.json$/i.exec(x.name || ""); if(!m) return null;
      var u = (x.createdBy && x.createdBy.user) || {};
      return { obra: m[1], em: x.createdDateTime || null, email: u.email || null, nome: u.displayName || null, arquivo: x.name };
    }).filter(function(x){ return x && x.email; });
    gravarLS("trocas:" + manifestoCaminho, { dados: t, em: new Date().toISOString() });
    return { dados: t, origem: "sharepoint" };
  }).catch(function(e){
    if(e && e.status === 404) return { dados: [], origem: "nao_existe" };
    var c = lerLS("trocas:" + manifestoCaminho, null);
    return c ? { dados: c.dados, origem: "cache", em: c.em } : { dados: [], origem: "erro", erro: String(e && e.message || e) };
  });
}
/* ------------------------------------------------ ajustes de histórico feitos pelo ADM (Gustavo, 19/11)
   historico/<obraId>__<AAAAMMDD-HHMMSS>.json  {versao, acao:"adicionar"|"remover", obra, email, nome, de, ate}
   Um arquivo por ajuste, nunca sobrescrito (mesmo esquema das trocas). Só vale o que foi CRIADO por um ADM da
   lista de acessos ("Criado por" do SharePoint). Para cada (obra, líder) vale o ajuste mais recente. */
function pastaHistorico(manifestoCaminho){ return aoLado(manifestoCaminho, "historico"); }
/* listar(pasta) e baixar(url) -> Promise<texto> */
function carregarAjustes(listar, baixar, manifestoCaminho){
  var chave = "ajustes:" + manifestoCaminho;
  return listar(pastaHistorico(manifestoCaminho)).then(function(itens){
    var lst = (itens || []).filter(function(x){ return /^(.+)__(\d{8}-\d{6})(?:-\d+)?\.json$/i.test(x.name || ""); });
    return lst.reduce(function(ac, x){
      return ac.then(function(out){
        return baixar(x.url).then(function(txt){
          var d = JSON.parse(String(txt).replace(/^﻿/, "")), u = (x.createdBy && x.createdBy.user) || {};
          if(d && d.obra && d.email && (d.acao === "adicionar" || d.acao === "remover"))
            out.push({ acao: d.acao, obra: d.obra, email: d.email, nome: d.nome || null, de: d.de || null, ate: d.ate || null,
                       em: x.createdDateTime || d.em || null, arquivo: x.name, autorEmail: u.email || null, autorNome: u.displayName || null });
          return out;
        }, function(){ return out; });
      });
    }, Promise.resolve([])).then(function(out){ gravarLS(chave, { dados: out, em: new Date().toISOString() }); return { dados: out, origem: "sharepoint" }; });
  }).catch(function(e){
    if(e && e.status === 404) return { dados: [], origem: "nao_existe" };
    var c = lerLS(chave, null);
    return c ? { dados: c.dados, origem: "cache", em: c.em } : { dados: [], origem: "erro", erro: String(e && e.message || e) };
  });
}
function novoAjuste(acao, obra, lider, de, ate, adm, agora){
  agora = agora || new Date();
  function d2(n){ return (n < 10 ? "0" : "") + n; }
  var carimbo = agora.getFullYear() + d2(agora.getMonth() + 1) + d2(agora.getDate()) + "-" + d2(agora.getHours()) + d2(agora.getMinutes()) + d2(agora.getSeconds());
  return { nome: obra.id + "__" + carimbo + ".json",
           conteudo: JSON.stringify({ versao: 1, acao: acao, obra: obra.id, nomeObra: obra.nome, email: lider.email, nome: lider.nome,
                                      de: de || null, ate: ate || null, por: adm ? adm.email : null, em: agora.toISOString() }) };
}
function ehAdm(email, acessos){ return ((acessos && acessos.adms) || []).some(function(x){ return mesmoEmail(x.email, email); }); }
/* aplica os ajustes válidos às "realizadas" de uma conta */
function aplicarAjustes(realizadas, manifesto, p, ctx, idsMinhas){
  var porObra = {};
  (ctx.ajustes || []).filter(function(a){ return mesmoEmail(a.email, p.email) && ehAdm(a.autorEmail, ctx.acessos); })
    .sort(function(a, b){ return String(a.em).localeCompare(String(b.em)) || String(a.arquivo).localeCompare(String(b.arquivo)); })
    .forEach(function(a){ porObra[a.obra] = a; });
  Object.keys(porObra).forEach(function(id){
    var a = porObra[id];
    realizadas = realizadas.filter(function(h){ return h.obra.id !== id; });
    if(a.acao !== "adicionar") return;
    var ativa = ((manifesto && manifesto.obras) || []).filter(function(o){ return o.id === id; })[0];
    var enc = ((manifesto && manifesto.encerradas) || []).filter(function(o){ return o.id === id; })[0];
    var o = ativa || enc; if(!o || (ativa && idsMinhas[id] && !a.ate)) return;
    realizadas.push({ obra: o, ativa: !!ativa, papel: "editor", de: a.de, ate: a.ate, confirmado: true, manual: true, ajuste: a });
  });
  return realizadas;
}
/* para o ADM: o que cada líder da lista enxerga (em andamento e histórico) */
function historicoDosLideres(manifesto, ctx){
  ctx = contexto(ctx);
  return ((ctx.acessos && ctx.acessos.lideres) || []).map(function(l){
    var p = { nome: l.nome, email: l.email, adm: false, lider: true, papel: "lider" };
    var ls = listasDaConta(manifesto, p, ctx);
    return { lider: l, minhas: ls.minhas, realizadas: ls.realizadas };
  });
}
/* arquivo de uma troca nova: o app sobe em <pastaTrocas>/<nome>, sem sobrescrever */
function novaTroca(o, p, anterior, agora){
  agora = agora || new Date();
  function d2(n){ return (n < 10 ? "0" : "") + n; }
  var carimbo = agora.getFullYear() + d2(agora.getMonth() + 1) + d2(agora.getDate()) + "-" + d2(agora.getHours()) + d2(agora.getMinutes()) + d2(agora.getSeconds());
  return { nome: o.id + "__" + carimbo + ".json",
           conteudo: JSON.stringify({ versao: 1, obra: o.id, nomeObra: o.nome, nome: p.nome, email: p.email, em: agora.toISOString(),
                                      anterior: anterior ? { nome: anterior.nome, email: anterior.email || null } : null }) };
}

/* pessoa da lista de acessos pelo nome escrito no obras.json ("Fábio", "Alves"):
   igual ao primeiro nome, a um sobrenome ou a uma parte do e-mail. Só vale se achar UMA pessoa. */
function pessoasDaLista(acessos){ return [].concat((acessos && acessos.lideres) || [], (acessos && acessos.adms) || []); }
function resolverNome(nome, acessos){
  var n = norm(nome); if(!n) return null;
  var achou = pessoasDaLista(acessos).filter(function(x){
    var partes = norm(x.nome).split(" ").concat(norm(String(x.email || "").split("@")[0]).split(/[._-]/));
    return partes.indexOf(n) >= 0 || norm(x.nome) === n;
  });
  return achou.length === 1 ? achou[0] : null;
}
/* origem = de onde veio o acessos.json: "sharepoint" | "cache" | "nao_existe" | "erro".
   Só a lista que NÃO EXISTE (404) volta ao modo da 0.2 (papel null). Erro sem cópia guardada = "indeterminado":
   só consulta, nada de marcar nem trocar, e a tela avisa. */
function perfilDaConta(c, acessos, origem){
  var email = c ? (c.username || "") : "", nome = c ? (c.name || "") : "";
  var adm = ((acessos && acessos.adms) || []).some(function(x){ return mesmoEmail(x.email, email); });
  var lider = ((acessos && acessos.lideres) || []).some(function(x){ return mesmoEmail(x.email, email); });
  var papel = acessos ? (adm ? "adm" : lider ? "lider" : "visualizador") : (origem === "erro" ? "indeterminado" : null);
  return { nome: nome, email: email, adm: adm, lider: lider, papel: papel };
}

/* ------------------------------------------------ listas da conta (tela inicial e histórico)
   minhas: obras ativas em que a conta é editora;
   outras: as demais obras ativas (todo mundo vê todo mundo: decisão do Gustavo, 04/11);
   realizadas: encerradas + ativas em que a conta JÁ FOI editora ou escalada.
   Cada item diz se a ligação com a conta está confirmada (e-mail) ou é por nome. */
function registroDaObra(editores, id){ return (editores && editores.obras && editores.obras[id]) || null; }
/* em ordem: data de criação no SharePoint e, no empate (mesmo segundo), o nome do arquivo. A última vale. */
function trocasDaObra(trocas, id){
  return (trocas || []).filter(function(t){ return t.obra === id; })
    .sort(function(a, b){ return String(a.em).localeCompare(String(b.em)) || String(a.arquivo).localeCompare(String(b.arquivo)); });
}
function papelNoHistorico(hist, p){
  var out = [];
  (hist || []).forEach(function(h){
    if(mesmoEmail(h.email, p.email)) out.push({ papel: h.papel || "editor", de: h.de || null, ate: h.ate || null, confirmado: true });
    else if(!h.email && mesmoNome(h.nome, p.nome)) out.push({ papel: h.papel || "editor", de: h.de || null, ate: h.ate || null, confirmado: false });
  });
  return out;
}
/* ctx = {editores, trocas, acessos} (aceita também só o editores, como na primeira 0.3) */
function contexto(x){ return x && (x.trocas || x.acessos || x.editores || x.ajustes) ? x : { editores: x || null, trocas: [], acessos: null, ajustes: [] }; }
function editorDaObra(o, ctx){
  ctx = contexto(ctx);
  var tr = trocasDaObra(ctx.trocas, o.id), ult = tr[tr.length - 1];
  if(ult) return { nome: ult.nome, email: ult.email, desde: ult.em, confirmado: true, fonte: "troca" };
  var r = registroDaObra(ctx.editores, o.id);
  if(r && r.editor) return { nome: r.editor.nome, email: r.editor.email || null, desde: r.editor.desde || null, confirmado: !!r.editor.email, fonte: "registro" };
  if(o.lider){
    var pessoa = resolverNome(o.lider, ctx.acessos);
    if(pessoa) return { nome: pessoa.nome, email: pessoa.email, desde: null, confirmado: true, fonte: "obras.json+acessos" };
    return { nome: o.lider, email: null, desde: null, confirmado: false, fonte: "obras.json", naoIdentificado: !!ctx.acessos };
  }
  return null;
}
function ehEditor(o, p, ctx){
  var ed = editorDaObra(o, ctx);
  if(!ed) return false;
  if(ed.email) return mesmoEmail(ed.email, p.email);
  if(ed.naoIdentificado || p.papel) return false;     /* com lista de acessos, só e-mail decide */
  return algumNome(ed.nome, p.nome);                  /* sem lista: como a 0.2 */
}
/* com a lista de acessos: só líder marca (adm e visualizador consultam) */
/* Líder que veio de uma TROCA: a planta passa a enviar com o primeiro nome da conta (o nome vai no pacote e
   no nome do arquivo). Grava no que a planta guarda no aparelho ANTES de ela abrir; a planta lê na partida.
   O líder original da obra fica como sempre foi (o nome que veio na planta). Devolve o nome gravado ou null. */
function nomeDoLiderNaPlanta(o, p, ctx, info){
  ctx = contexto(ctx);
  if(!info || !info.chave) return null;
  var ed = editorDaObra(o, ctx);
  if(!ed || ed.fonte !== "troca" || !mesmoEmail(ed.email, p.email)) return null;
  var nome = norm(p.nome).split(" ")[0].replace(/[^a-z0-9]/g, "");
  if(!nome) return null;
  nome = nome.charAt(0).toUpperCase() + nome.slice(1);
  try{
    var d = null; try{ d = JSON.parse(localStorage.getItem(info.chave) || "null"); }catch(e){}
    if(!d || typeof d !== "object") d = { v: 1, lider: "", reg: {}, ocultos: {} };
    if(d.lider === nome) return nome;
    d.lider = nome; localStorage.setItem(info.chave, JSON.stringify(d));
  }catch(e){ return null; }
  return nome;
}
function podeMarcar(p){ return p.papel === null || p.papel === undefined || p.papel === "lider"; }
/* -> Promise<bool>: lê o que a planta guarda neste aparelho (precisa da CHAVE/banco: guardados ao abrir a obra) */
function pendenciaDoAparelho(id){
  var info = lerLS("info:" + id, null);
  if(!info) return Promise.resolve(temPendencia(id));
  return lerEstadoPlanta(info).then(function(e){
    if(!e) return temPendencia(id);
    if(e.fotosPendentes) return true;
    if(!e.marcadas) return false;
    var s = estadoGuardado(id);
    if(s.confirmado && s.confirmado.dados) return s.confirmado.dados !== e.dados;
    return true;   /* sem envio confirmado pela 0.3, marcação no aparelho conta como pendente (revisão 19/11, X3) */
  }, function(){ return temPendencia(id); });
}
function lembrarPlanta(id, info){ if(info) gravarLS("info:" + id, { chave: info.chave, banco: info.banco }); }
/* há marcação ou foto deste aparelho que ainda não subiu? (pelo que a casca já viu da planta: online:envio:<id>) */
function temPendencia(id){
  var s = estadoGuardado(id), l = s.lido;
  if(!l) return false;
  if(l.fotosPendentes) return true;
  if(!l.marcadas) return false;
  return !s.confirmado || !s.confirmado.dados || s.confirmado.dados !== l.dados;
}
/* "Agora sou o líder": líder da lista, obra que ainda não é dele. Sem lista: o atalho antigo, só quando o
   responsável não está confirmado por e-mail. */
/* "Agora sou o líder": só líder da LISTA de acessos. Sem lista (ou lista que não deu para ler) ninguém troca:
   uma troca gravada antes da lista existir continuaria valendo depois (achado da validação de 18/11). */
function podeAssumir(o, p, ctx){
  ctx = contexto(ctx);
  if(o.somenteConsulta || p.papel !== "lider") return false;
  return !ehEditor(o, p, ctx);
}
/* histórico que sai das trocas: o líder do obras.json até a 1ª troca, e cada troca até a seguinte */
/* período em aberto do registro termina na primeira troca que veio depois dele */
function fecharPeloRegistro(hist, o, ctx){
  var tr = trocasDaObra(ctx.trocas, o.id);
  return (hist || []).map(function(h){
    if(h.ate || !tr.length) return h;
    var depois = tr.filter(function(t){ return !h.de || t.em.slice(0, 10) >= h.de; })[0];
    if(!depois || (depois.email && mesmoEmail(depois.email, h.email))) return h;
    var c = {}; for(var k in h) c[k] = h[k]; c.ate = depois.em.slice(0, 10); return c;
  });
}
function historicoDasTrocas(o, ctx){
  var tr = trocasDaObra(ctx.trocas, o.id); if(!tr.length) return [];
  var h = [], ini = o.lider ? (resolverNome(o.lider, ctx.acessos) || { nome: o.lider }) : null;
  if(ini) h.push({ nome: ini.nome, email: ini.email || null, de: null, ate: tr[0].em.slice(0, 10), papel: "escalado" });
  tr.forEach(function(t, i){ h.push({ nome: t.nome, email: t.email, de: t.em.slice(0, 10), ate: tr[i + 1] ? tr[i + 1].em.slice(0, 10) : null, papel: "editor" }); });
  return h;
}
function listasDaConta(manifesto, p, ctx){
  ctx = contexto(ctx);
  var minhas = [], outras = [], realizadas = [];
  ((manifesto && manifesto.obras) || []).forEach(function(o){
    var ed = editorDaObra(o, ctx), meu = podeMarcar(p) && !o.somenteConsulta && ehEditor(o, p, ctx);
    (meu ? minhas : outras).push({ obra: o, editor: ed, confirmado: !!(ed && ed.confirmado), podeAssumir: podeAssumir(o, p, ctx) });
    var r = registroDaObra(ctx.editores, o.id);
    var hr = fecharPeloRegistro((r && r.historico) || [], o, ctx), ht = historicoDasTrocas(o, ctx);
    if(hr.length && ht.length && ht[0].papel === "escalado") ht = ht.slice(1);   /* o registro já diz quem estava antes */
    papelNoHistorico(hr.concat(ht), p).forEach(function(h){
      if(meu && !h.ate) return;                          /* período em aberto = é a obra atual, já está em "minhas" */
      realizadas.push({ obra: o, ativa: true, papel: h.papel, de: h.de, ate: h.ate, confirmado: h.confirmado });
    });
  });
  ((manifesto && manifesto.encerradas) || []).forEach(function(o){
    if(p.adm){   /* Gustavo, 18/11: ADM (ele e o Ighor) vê todo o histórico, não só onde trabalhou */
      var nomes = ((o.lideres || []).map(function(l){ return l.nome; }).filter(Boolean));
      if(!nomes.length && o.lider) nomes = [o.lider];
      realizadas.push({ obra: o, ativa: false, papel: "adm", de: o.inicio || null, ate: o.encerrada_em || null, confirmado: true, lideres: nomes });
      return;
    }
    var r = registroDaObra(ctx.editores, o.id), hist = (r && r.historico) || o.lideres || [];
    hist = hist.concat(historicoDasTrocas(o, ctx));
    var ps = papelNoHistorico(hist, p);
    if(!ps.length && !hist.length && o.lider){
      var pessoa = resolverNome(o.lider, ctx.acessos);
      if(pessoa ? mesmoEmail(pessoa.email, p.email) : mesmoNome(o.lider, p.nome))
        ps = [{ papel: "editor", de: null, ate: o.encerrada_em || null, confirmado: !!pessoa }];
    }
    ps.forEach(function(h){ realizadas.push({ obra: o, ativa: false, papel: h.papel, de: h.de, ate: h.ate, confirmado: h.confirmado }); });
  });
  var idsMinhas = {}; minhas.forEach(function(it){ idsMinhas[it.obra.id] = true; });
  if(!p.adm) realizadas = aplicarAjustes(realizadas, manifesto, p, ctx, idsMinhas);
  realizadas.sort(function(a, b){ return String(b.ate || b.de || "9").localeCompare(String(a.ate || a.de || "9")); });
  return { minhas: minhas, outras: outras, realizadas: realizadas };
}

/* ------------------------------------------------ modo da obra
   "editor"       : abre a planta normal (é o comportamento da 0.2);
   "leitura"      : a conta não é a editora registrada (a tela de leitura vem do bloco D);
   "indisponivel" : leitura pedida numa planta sem adaptador (por módulo / remanejamento: pendente). */
var FORMATOS_COM_ADAPTADOR = ["planta-lider"];
function formatoDaPlanta(html){
  var m = /formato\s*:\s*"([^"]+)"/.exec(html || "");
  return m ? m[1] : null;
}
function modoDaObra(o, p, ctx, htmlPlanta){
  ctx = contexto(ctx);
  var r = registroDaObra(ctx.editores, o.id), tem = r && r.editor || trocasDaObra(ctx.trocas, o.id).length || ctx.acessos;
  if(!tem && !o.somenteConsulta) return { modo: "editor", motivo: "sem registro de editor: abre como hoje" };
  if(!o.somenteConsulta && podeMarcar(p) && ehEditor(o, p, ctx)) return { modo: "editor", motivo: "conta é a editora" };
  var fmt = formatoDaPlanta(htmlPlanta);
  if(FORMATOS_COM_ADAPTADOR.indexOf(fmt) < 0) return { modo: "indisponivel", motivo: "visualização indisponível para esta planta (" + (fmt || "formato desconhecido") + ")" };
  var ed = editorDaObra(o, ctx);
  return { modo: "leitura", motivo: "editor: " + (ed ? (ed.nome || ed.email) : "não definido") };
}

/* ------------------------------------------------ estado do envio (faixa)
   Guardado por obra em online:envio:<id> = {
     pendente:   {em, arquivos:[nome], parte, partes}           envio começou e ainda não confirmou tudo
     confirmado: {em, arquivos:[nome], parte, partes, completo, exportadoEm, sig}
     alteracao:  {sig, em}                                        última mudança vista na planta
   }
   Sobre o "pacote pendente" sugerido na revisão: a planta 1.03u/1.04 só marca fotos como enviadas
   depois da confirmação de cada parte; o que não confirmou continua pendente DENTRO da planta e sai
   no próximo ENVIAR. Por isso o núcleo guarda o REGISTRO do envio interrompido (para o vermelho
   voltar ao reabrir), não uma segunda cópia do pacote (até 25 MB por parte no tablet). */
var ENVIANDO_OBRA = {};
function estadoGuardado(id){ return lerLS("envio:" + id, {}); }
function guardarEstado(id, s){ gravarLS("envio:" + id, s); }

/* lê parte/partes/exportadoEm do pacote sem carregar as fotos (vêm no fim do arquivo) */
function cabecalhoDoPacote(f){
  if(!f || !/\.txt$/i.test(f.name)) return Promise.resolve(null);
  var lim = Math.min(f.size, 4 * 1024 * 1024);
  return f.slice(0, lim).text().then(function(t){
    function num(k){ var m = new RegExp('"' + k + '"\\s*:\\s*(\\d+)').exec(t); return m ? +m[1] : null; }
    var ex = /"exportadoEm"\s*:\s*"([^"]+)"/.exec(t), fm = /"formato"\s*:\s*"([^"]+)"/.exec(t);
    return { formato: fm ? fm[1] : null, exportadoEm: ex ? ex[1] : null, parte: num("parte"), partes: num("partes") };
  }).catch(function(){ return null; });
}

/* adaptador da planta planta-lider (1.03u/1.04).
   A planta roda dentro de uma função fechada: assinatura(), PEND e faltaParte() NÃO são visíveis de fora
   (conferido no laboratório, 18/11). Por isso o adaptador lê o que a planta GRAVA no aparelho, na mesma
   origem da casca: localStorage[CHAVE] (marcações), CHAVE+":oc" (ocorrências) e o banco de fotos
   (fotos com enviada=false). Uma planta futura que avise a casca de cada alteração é mais robusta. */
function infoDaPlanta(html){
  var fmt = formatoDaPlanta(html), mc = /var CHAVE\s*=\s*"([^"]+)"/.exec(html || ""), b = bancoDeFotos(html);
  if(FORMATOS_COM_ADAPTADOR.indexOf(fmt) < 0 || !mc || !b) return null;
  return { formato: fmt, chave: mc[1], banco: b };
}
function hashTexto(s){ var h1 = 0x811c9dc5, h2 = 7; for(var i = 0; i < s.length; i++){ h1 = Math.imul(h1 ^ s.charCodeAt(i), 16777619); h2 = (Math.imul(h2, 31) + s.charCodeAt(i)) | 0; } return (h1 >>> 0).toString(16) + (h2 >>> 0).toString(16) + ":" + s.length; }
function fotosNaoEnviadas(banco){
  return new Promise(function(ok){
    var lista = indexedDB.databases ? indexedDB.databases() : Promise.resolve(null);
    lista.then(function(ds){
      if(ds && !ds.some(function(d){ return d.name === banco; })) return ok([]);
      var r = indexedDB.open(banco);
      r.onsuccess = function(){ var db = r.result;
        if(!db.objectStoreNames.contains("fotos")){ db.close(); return ok([]); }
        var ids = [], c = db.transaction("fotos").objectStore("fotos").openCursor();
        c.onsuccess = function(e){ var cur = e.target.result; if(!cur){ db.close(); return ok(ids); }
          if(!cur.value.enviada) ids.push(cur.value.id); cur.continue(); };
        c.onerror = function(){ db.close(); ok(null); }; };
      r.onerror = function(){ ok(null); };
    }, function(){ ok(null); });
  });
}
/* Conteúdo das marcações (reg + ocorrências) em forma canônica: chaves em ordem, para que o que
   está no tablet e o que está DENTRO do arquivo enviado deem o mesmo resumo quando são iguais. */
function canonico(v){
  if(v === null || typeof v !== "object") return JSON.stringify(v === undefined ? null : v);
  if(Array.isArray(v)) return "[" + v.map(canonico).join(",") + "]";
  return "{" + Object.keys(v).sort().map(function(k){ return JSON.stringify(k) + ":" + canonico(v[k]); }).join(",") + "}";
}
function semMarcaDeEnvio(oc){
  if(!oc || typeof oc !== "object") return oc;
  function limpa(x){ if(!x || typeof x !== "object") return x; var c = {}; for(var k in x) if(k !== "enviada" && k !== "enviadaEm") c[k] = x[k]; return c; }
  if(Array.isArray(oc)) return oc.map(limpa);
  var m = {}; for(var k in oc) m[k] = limpa(oc[k]); return m;
}
/* a planta marca cada ocorrência como enviada DEPOIS do envio (marcarOcEnviadas): isso não é alteração */
function resumoDados(reg, oc){ return hashTexto(canonico(reg || {}) + "|" + canonico(semMarcaDeEnvio(oc) || [])); }
/* -> Promise<{sig, dados, fotosPendentes}> ou null (planta sem adaptador)
   sig   = tudo o que mudou no tablet (para "última alteração");
   dados = só reg + ocorrências, comparável com o que veio no arquivo enviado. */
function lerEstadoPlanta(info){
  if(!info) return Promise.resolve(null);
  var reg = "", oc = "";
  try{ reg = localStorage.getItem(info.chave) || ""; oc = localStorage.getItem(info.chave + ":oc") || ""; }catch(e){}
  var r = null, o = null;
  try{ var j = JSON.parse(reg || "null"); r = j && typeof j === "object" && "reg" in j ? j.reg : j; }catch(e){}
  try{ o = JSON.parse(oc || "null"); }catch(e){}
  var marcadas = 0, ultimaMarca = "";
  if(r && typeof r === "object") Object.keys(r).forEach(function(k){ var x = r[k] || {}, st = x.st; if(st && st !== "pendente") marcadas++;
    [x.ini, x.fim].forEach(function(d){ if(d && String(d) > ultimaMarca) ultimaMarca = String(d); }); });
  if(Array.isArray(o)) marcadas += o.length;
  else if(o && typeof o === "object") Object.keys(o).forEach(function(k){ if(o[k] && o[k].txt && !o[k].apagada) marcadas++; });
  var okEnvio = null; try{ okEnvio = localStorage.getItem(info.chave + ":okEnvio"); }catch(e){}
  return fotosNaoEnviadas(info.banco).then(function(ids){
    ids = ids || [];
    return { sig: hashTexto(reg + "|" + oc + "|" + ids.join(",")), dados: resumoDados(r, o), fotosPendentes: ids.length, marcadas: marcadas,
             okEnvio: okEnvio, ultimaMarca: ultimaMarca || null };
  });
}

/* Lote = um ENVIAR da planta (1 arquivo, ou N partes com o mesmo carimbo no nome).
   O verde depende do que REALMENTE subiu: o núcleo lê o começo de cada arquivo enviado (reg e oc vêm
   antes das fotos) e guarda o resumo desse conteúdo. Alteração feita durante o upload não está no
   arquivo, então não bate com o tablet e a faixa fica amarela até o próximo envio confirmado (pedido do
   Codex, 18/11). Envio em partes só fica completo quando TODAS as partes do mesmo lote confirmaram:
   receber a última parte não basta. */
var CABECAS_ENVIO = {}, EM_VOO = {};
function lerArquivoEnviado(f){
  if(!f || !/\.txt$/i.test(f.name)) return Promise.resolve(null);
  return f.slice(0, Math.min(f.size, CABECA)).text().then(function(t){
    var c = lerCabeca(t);
    return { formato: c.formato, exportadoEm: c.exportadoEm, parte: c.parte, partes: c.partes,
             dados: c.reg === undefined ? null : resumoDados(c.reg, c.oc), lote: loteDoArquivo(f.name) };
  }).catch(function(){ return null; });
}
function pendentesDe(s){
  var m = s.pendentes && typeof s.pendentes === "object" ? s.pendentes : {};
  if(s.pendente && s.pendente.arquivos) m[s.pendente.arquivos[0] || "antigo"] = s.pendente;   /* registro da 1ª 0.3 */
  return m;
}
function envioComecou(o, arqs){
  ENVIANDO_OBRA[o.id] = (ENVIANDO_OBRA[o.id] || 0) + 1;
  var s = estadoGuardado(o.id), nome = arqs[0] && arqs[0].name;
  s.pendentes = pendentesDe(s); delete s.pendente;
  s.pendentes[nome] = { em: new Date().toISOString(), arquivos: arqs.map(function(f){ return f.name; }), parte: null, partes: null };
  EM_VOO[o.id + "|" + nome] = true;
  guardarEstado(o.id, s);
  var p = lerArquivoEnviado(arqs[0]);
  CABECAS_ENVIO[o.id + "|" + nome] = p;
  return p.then(function(c){
    if(!c) return;
    var s2 = estadoGuardado(o.id); s2.pendentes = pendentesDe(s2);
    var pe = s2.pendentes[nome];
    if(pe){ pe.parte = c.parte; pe.partes = c.partes; pe.exportadoEm = c.exportadoEm; guardarEstado(o.id, s2); }
  });
}
/* -> Promise (resolve quando o registro do envio estiver gravado) */
function envioTerminou(o, arqs, ok, info){
  function fim(){ ENVIANDO_OBRA[o.id] = Math.max(0, (ENVIANDO_OBRA[o.id] || 1) - 1); delete EM_VOO[o.id + "|" + (arqs[0] && arqs[0].name)]; }
  var chave = o.id + "|" + (arqs[0] && arqs[0].name);
  var cab = CABECAS_ENVIO[chave] || lerArquivoEnviado(arqs[0]);
  delete CABECAS_ENVIO[chave];
  if(!ok){ fim(); return Promise.resolve(); }   /* falhou: o pendente fica (vermelho) até um envio confirmar */
  var em = new Date().toISOString();
  return cab.then(function(c){
    c = c || {};
    var s = estadoGuardado(o.id), partes = c.partes || null, parte = c.parte || null;
    /* partes recebidas do mesmo lote; lote novo recomeça a conta */
    var lote = s.lote && s.lote.id === c.lote && c.lote ? s.lote : { id: c.lote || null, partes: partes, recebidas: [] };
    if(parte && lote.recebidas.indexOf(parte) < 0) lote.recebidas.push(parte);
    var todas = !partes || partes === 1;
    if(!todas){ todas = true; for(var k = 1; k <= partes; k++) if(lote.recebidas.indexOf(k) < 0){ todas = false; break; } }
    s.lote = lote;
    s.confirmado = { em: em, arquivos: arqs.map(function(f){ return f.name; }), parte: parte, partes: partes,
                     recebidas: lote.recebidas.slice().sort(), exportadoEm: c.exportadoEm || null,
                     completo: todas, dados: c.dados || null };
    s.pendentes = pendentesDe(s); delete s.pendente;
    delete s.pendentes[arqs[0] && arqs[0].name];
    if(todas) Object.keys(s.pendentes).forEach(function(n){ if(!EM_VOO[o.id + "|" + n] && s.pendentes[n].em <= em) delete s.pendentes[n]; });
    guardarEstado(o.id, s);
  }).then(fim, fim);
}
/* chamada pela casca (ao abrir, a cada 30 s, ao voltar ao app): lê o estado da planta e registra a
   hora da última alteração vista. Devolve Promise. */
function observar(o, info){
  return lerEstadoPlanta(info).then(function(e){
    if(!e) return null;
    var s = estadoGuardado(o.id), agora = new Date().toISOString();
    if(!s.alteracao) s.alteracao = { sig: e.sig, em: e.ultimaMarca || (s.confirmado && s.confirmado.em) || agora };
    else if(s.alteracao.sig !== e.sig) s.alteracao = { sig: e.sig, em: agora };
    s.lido = { sig: e.sig, dados: e.dados, fotosPendentes: e.fotosPendentes, marcadas: e.marcadas, okEnvio: e.okEnvio || null, em: agora };
    guardarEstado(o.id, s);
    return e;
  });
}
/* {cor, salvoNoTablet, confirmadoNoSharePoint, ultimaAlteracao, partes, faltam:{fotos, partes}, motivos[]}
   Verde só quando: o último lote confirmou TODAS as partes, o conteúdo que subiu é igual ao do tablet
   e não há foto pendente. Usa a última leitura do observar(): chamar observar() antes. */
function estadoEnvio(o, info){
  var s = estadoGuardado(o.id), mot = [];
  var r = { cor: null, salvoNoTablet: true, confirmadoNoSharePoint: null,
            ultimaAlteracao: s.alteracao ? s.alteracao.em : null,
            partes: s.confirmado ? { parte: s.confirmado.parte, de: s.confirmado.partes, recebidas: s.confirmado.recebidas || [] } : null,
            faltam: { fotos: null, partes: null }, motivos: mot, adaptador: !!info };
  /* "confirmado no SharePoint" só com o lote inteiro: a parte 2 de 3 não vira data de "Tudo enviado" */
  if(s.confirmado && s.confirmado.completo) r.confirmadoNoSharePoint = s.confirmado.em;
  if(ENVIANDO_OBRA[o.id] > 0){ r.cor = "azul"; mot.push("enviando"); return r; }
  var pend = pendentesDe(s), pk = Object.keys(pend);
  if(pk.length){ r.cor = "vermelho"; mot.push("envio de " + pend[pk[0]].em + " começou e não confirmou"); return r; }
  if(!info || !s.lido){ r.cor = s.confirmado ? "desconhecido" : "nada"; mot.push(info ? "estado da planta ainda não lido" : "planta sem adaptador: só se sabe o que este aparelho enviou"); return r; }
  r.faltam.fotos = s.lido.fotosPendentes;
  /* nada marcado, nenhuma foto e nada enviado deste aparelho: não há o que avisar (faixa neutra) */
  if(!s.confirmado && !s.lido.marcadas && !s.lido.fotosPendentes){ r.cor = "vazio"; mot.push("nada marcado ainda"); return r; }
  /* veio da 0.2: a planta registra um envio, mas a 0.3 ainda não confirmou nenhum. Não dá para provar que o que está
     no tablet é igual ao que subiu (ex.: ocorrência feita depois). Neutro, pedindo um ENVIAR (revisão 19/11, X3). */
  if(!s.confirmado && s.lido.okEnvio){ r.cor = "migracao"; r.envioAnterior = s.lido.okEnvio; mot.push("último envio foi pela versão anterior"); return r; }
  if(!s.confirmado){ r.cor = "amarelo"; mot.push("nada confirmado neste aparelho"); return r; }
  if(!s.confirmado.completo){
    var falta = []; for(var k = 1; k <= (s.confirmado.partes || 0); k++) if((s.confirmado.recebidas || []).indexOf(k) < 0) falta.push(k);
    r.faltam.partes = falta.length; mot.push("faltam partes do envio: " + falta.join(", ") + " de " + s.confirmado.partes);
  }
  if(s.lido.fotosPendentes) mot.push(s.lido.fotosPendentes + " foto(s) não enviada(s)");
  if(!s.confirmado.dados) mot.push("conteúdo do envio não conferido");
  else if(s.confirmado.dados !== s.lido.dados) mot.push("alterações que não estão no último envio");
  r.cor = mot.length ? "amarelo" : "verde";
  return r;
}

/* ------------------------------------------------ espaço no tablet
   Só mede. Apagar foto fica para quando der para provar, foto a foto, que há cópia no SharePoint
   (decisão do Gustavo, 18/11). Não abre banco que não existe (indexedDB.open criaria um vazio). */
function bancoDeFotos(html){ var m = /indexedDB\.open\(\s*"([^"]+)"/.exec(html || ""); return m ? m[1] : null; }
function somarBanco(nome){
  return new Promise(function(ok){
    var r = indexedDB.open(nome);
    r.onsuccess = function(){
      var db = r.result;
      if(!db.objectStoreNames.contains("fotos")){ db.close(); return ok({ fotos: 0, bytes: 0, naoEnviadas: 0 }); }
      var t = { fotos: 0, bytes: 0, naoEnviadas: 0 }, c = db.transaction("fotos").objectStore("fotos").openCursor();
      c.onsuccess = function(e){ var cur = e.target.result;
        if(!cur){ db.close(); return ok(t); }
        var v = cur.value; t.fotos++; t.bytes += (v.blob && v.blob.size) || 0; if(!v.enviada) t.naoEnviadas++;
        cur.continue(); };
      c.onerror = function(){ db.close(); ok(null); };
    };
    r.onerror = function(){ ok(null); };
  });
}
/* obras: [{id, nome, html}] (html = planta guardada no cache da casca) */
function espacoUsado(obras){
  var est = (navigator.storage && navigator.storage.estimate) ? navigator.storage.estimate() : Promise.resolve({});
  var lista = (indexedDB.databases ? indexedDB.databases() : Promise.resolve(null)).catch(function(){ return null; });
  return Promise.all([est, lista]).then(function(x){
    var e = x[0] || {}, existentes = x[1] ? x[1].map(function(d){ return d.name; }) : null;
    return Promise.all((obras || []).map(function(o){
      var b = bancoDeFotos(o.html);
      if(!b) return { id: o.id, nome: o.nome, banco: null, fotos: null, bytes: null, naoEnviadas: null };
      if(existentes && existentes.indexOf(b) < 0) return { id: o.id, nome: o.nome, banco: b, fotos: 0, bytes: 0, naoEnviadas: 0 };
      return somarBanco(b).then(function(t){ return { id: o.id, nome: o.nome, banco: b, fotos: t && t.fotos, bytes: t && t.bytes, naoEnviadas: t && t.naoEnviadas }; });
    })).then(function(por){
      return { usado: e.usage || null, cota: e.quota || null, porObra: por, medidoEm: new Date().toISOString() };
    });
  });
}

/* ------------------------------------------------ consulta: a mesma planta em outro aparelho
   Prioridade da 0.3 (Gustavo, 18/11): o Ighor abre a obra do Fábio e vê as marcações do ÚLTIMO ENVIO
   COMPLETO CONFIRMADO, com autor e horário. O que ainda está só no tablet do Fábio entra depois do envio.

   - o envio é escolhido pela LISTA da pasta Envios (data de criação e "Criado por" do SharePoint),
     não pelo nome do arquivo; envio em partes só conta se todas as partes estão lá;
   - do arquivo só se baixa o começo (Range): reg e oc vêm antes das fotos. Fotos ficam para a 0.4;
   - a planta abre ISOLADA: CHAVE e banco de fotos ganham o sufixo ":consulta", então nada do que o
     aparelho já guarda (inclusive de quem já foi editor da obra) é lido ou alterado. */
var CABECA = 4 * 1024 * 1024;
/* pega o valor JSON que começa depois de "chave": num texto truncado (só objetos e listas) */
function valorJSON(t, chave){
  var i = t.indexOf('"' + chave + '"'); if(i < 0) return undefined;
  i = t.indexOf(":", i) + 1; while(/\s/.test(t[i])) i++;
  var abre = t[i], fecha = abre === "{" ? "}" : abre === "[" ? "]" : null;
  if(!fecha){ var m = /^(null|true|false|-?\d+(\.\d+)?|"(?:[^"\\]|\\.)*")/.exec(t.slice(i)); return m ? JSON.parse(m[1]) : undefined; }
  var nivel = 0, str = false, esc = false;
  for(var k = i; k < t.length; k++){
    var c = t[k];
    if(str){ if(esc) esc = false; else if(c === "\\") esc = true; else if(c === '"') str = false; continue; }
    if(c === '"') str = true; else if(c === abre) nivel++; else if(c === fecha){ nivel--; if(nivel === 0) return JSON.parse(t.slice(i, k + 1)); }
  }
  return undefined;   /* cortado antes de fechar: o começo baixado não bastou */
}
/* obra e projeto que a planta escreve em cada envio (pacoteEnvio): é assim que se sabe de qual planta é um arquivo */
function identidadeDaPlanta(html){
  var m = /formato:"planta-lider"[^}]{0,300}?obra:"((?:[^"\\]|\\.)*)"[^}]{0,80}?projeto:"((?:[^"\\]|\\.)*)"/.exec(html || "");
  return m ? { obra: JSON.parse('"' + m[1] + '"'), projeto: JSON.parse('"' + m[2] + '"') } : null;
}
function mesmaPlanta(c, id){ return !id || (c.obra === id.obra && c.projeto === id.projeto); }
function lerCabeca(texto){
  return { formato: valorJSON(texto, "formato") || null, versao: valorJSON(texto, "versao") || null,
           obra: valorJSON(texto, "obra") || null, projeto: valorJSON(texto, "projeto") || null,
           exportadoEm: valorJSON(texto, "exportadoEm") || null, lider: valorJSON(texto, "lider") || null,
           parte: valorJSON(texto, "parte") || null, partes: valorJSON(texto, "partes") || null,
           reg: valorJSON(texto, "reg"), oc: valorJSON(texto, "oc") };
}
/* lote de um arquivo de envio = o nome sem "_parteXdeN", sem "_Nfotos" e sem o "-2" de nome repetido.
   As partes de um mesmo ENVIAR têm o mesmo carimbo no nome (a planta monta todas com o mesmo "agora"). */
function loteDoArquivo(nome){
  return String(nome || "").replace(/(-\d{1,2})?\.(txt|json)$/i, "").replace(/_parte\d+de\d+.*$/i, "").replace(/_\d+fotos$/i, "");
}
function parteDoNome(nome){ var m = /_parte(\d+)de(\d+)/i.exec(nome || ""); return m ? { parte: +m[1], partes: +m[2] } : null; }
var CABECA_PEQUENA = 256 * 1024;
/* lê o começo do arquivo: 256 KB; se o reg/oc vierem cortados, pede 4 MB (reg e oc vêm antes das fotos) */
function lerInicio(x, deps){
  return deps.baixarInicio(x.url, Math.min(CABECA_PEQUENA, x.size || CABECA_PEQUENA)).then(function(t){
    var c = lerCabeca(t);
    if((c.reg === undefined || c.oc === undefined) && (x.size || 0) > CABECA_PEQUENA)
      return deps.baixarInicio(x.url, CABECA).then(function(t2){ return { c: lerCabeca(t2), texto: t2 }; });
    return { c: c, texto: t };
  });
}
/* o arquivo chegou inteiro? Pequeno: o JSON inteiro tem que abrir. Grande: o fim do arquivo tem que fechar o JSON. */
function arquivoInteiro(x, leitura, deps){
  var bytes = new TextEncoder().encode(leitura.texto).length;
  if((x.size || 0) && bytes >= x.size){ try{ JSON.parse(leitura.texto); return Promise.resolve(true); }catch(e){ return Promise.resolve(false); } }
  if(!deps.baixarFaixa || !x.size) return Promise.resolve(true);
  return deps.baixarFaixa(x.url, Math.max(0, x.size - 64), x.size - 1).then(function(fim){ return /\}\s*$/.test(fim); }, function(){ return false; });
}
/* deps = {listar(pasta) -> Promise<[{name,size,createdDateTime,createdBy,url}]>, baixarInicio(url, bytes),
           baixarFaixa(url, ini, fim) (opcional), identidade (obra/projeto da planta, opcional)}
   Procura do mais novo para o mais velho, na pasta inteira, o último LOTE completo e inteiro desta planta. */
function ultimoEnvioConfirmado(o, deps){
  return deps.listar(o.envios).then(function(itens){
    itens = (itens || []).filter(function(x){ return /\.(txt|json)$/i.test(x.name); })
      .sort(function(a, b){ return Date.parse(b.createdDateTime) - Date.parse(a.createdDateTime); });
    if(!itens.length) return { achou: false, motivo: "nenhum envio na pasta da obra", situacao: "sem_envio" };
    var lidos = {}, ignorados = [], vistos = {}, incompletoMaisNovo = null;
    function ler(x){
      if(lidos[x.name] !== undefined) return Promise.resolve(lidos[x.name]);
      return lerInicio(x, deps).then(function(l){ lidos[x.name] = l; return l; }, function(){ lidos[x.name] = null; return null; });
    }
    function valido(x, l){
      var c = l && l.c;
      if(!c || !c.reg){ ignorados.push({ arquivo: x.name, motivo: l ? "sem marcações legíveis" : "não deu para ler" }); return false; }
      if(FORMATOS_COM_ADAPTADOR.indexOf(c.formato) < 0){ ignorados.push({ arquivo: x.name, motivo: "formato " + c.formato + " sem consulta" }); return false; }
      if(!mesmaPlanta(c, deps.identidade)){ ignorados.push({ arquivo: x.name, motivo: "envio de outra planta (" + c.projeto + " · " + c.obra + ")" }); return false; }
      return true;
    }
    /* tenta o lote do arquivo i; devolve o resultado ou null para seguir procurando */
    function tentar(i){
      if(i >= itens.length) return Promise.resolve(null);
      var x = itens[i], lote = loteDoArquivo(x.name);
      if(vistos[lote]) return tentar(i + 1);
      vistos[lote] = true;
      return ler(x).then(function(l){
        if(!valido(x, l)) return tentar(i + 1);
        var c = l.c, partes = c.partes || 1, doLote = [x];
        if(partes > 1){
          doLote = itens.filter(function(y){ return loteDoArquivo(y.name) === lote; });
        }
        /* cada parte: lida, da mesma planta, do mesmo líder, e cada número de 1 a N uma vez */
        return doLote.reduce(function(ac, y){ return ac.then(function(acc){ return ler(y).then(function(ly){ acc.push({ x: y, l: ly }); return acc; }); }); }, Promise.resolve([]))
          .then(function(lst){
            var porParte = {};
            lst.forEach(function(e){ var cy = e.l && e.l.c;
              if(!cy || !cy.reg || !mesmaPlanta(cy, deps.identidade) || cy.lider !== c.lider || (cy.partes || 1) !== partes) return;
              var n = cy.parte || 1; if(!porParte[n] || Date.parse(e.x.createdDateTime) > Date.parse(porParte[n].x.createdDateTime)) porParte[n] = e; });
            var falta = []; for(var k = 1; k <= partes; k++) if(!porParte[k]) falta.push(k);
            if(falta.length){
              ignorados.push({ arquivo: x.name, motivo: "envio em partes incompleto (falta " + falta.join(", ") + " de " + partes + ")" });
              if(!incompletoMaisNovo) incompletoMaisNovo = { arquivo: x.name, em: x.createdDateTime, falta: falta, partes: partes };
              return tentar(i + 1);
            }
            var ordem = []; for(var k2 = 1; k2 <= partes; k2++) ordem.push(porParte[k2]);
            return ordem.reduce(function(ac, e){ return ac.then(function(ok){ return ok && arquivoInteiro(e.x, e.l, deps); }); }, Promise.resolve(true)).then(function(inteiro){
              if(!inteiro){ ignorados.push({ arquivo: x.name, motivo: "arquivo cortado (não chegou inteiro)" }); return tentar(i + 1); }
              var ult = ordem.reduce(function(a, e){ return Date.parse(e.x.createdDateTime) > Date.parse(a.x.createdDateTime) ? e : a; }, ordem[0]);
              var autor = ult.x.createdBy && ult.x.createdBy.user || {}, c1 = ordem[0].l.c;
              return { achou: true, situacao: "ok", arquivo: ult.x.name, arquivos: ordem.map(function(e){ return e.x.name; }), confirmadoEm: ult.x.createdDateTime,
                       autor: autor.displayName || null, autorEmail: autor.email || null,
                       lider: c1.lider, exportadoEm: c1.exportadoEm, formato: c1.formato, versao: c1.versao,
                       parte: partes > 1 ? partes : null, partes: partes > 1 ? partes : null, reg: c1.reg, oc: c1.oc || [],
                       incompletoMaisNovo: incompletoMaisNovo, ignorados: ignorados };
            });
          });
      });
    }
    return tentar(0).then(function(r){
      return r || { achou: false, situacao: incompletoMaisNovo ? "so_incompleto" : "sem_envio_valido",
                    motivo: "nenhum envio completo e inteiro desta planta na pasta", incompletoMaisNovo: incompletoMaisNovo, ignorados: ignorados };
    });
  });
}
/* HTML da planta isolada para consulta. Só para formato com adaptador; senão devolve null. */
function htmlDeConsulta(html, envio){
  if(FORMATOS_COM_ADAPTADOR.indexOf(formatoDaPlanta(html)) < 0) return null;
  var mc = /var CHAVE\s*=\s*"([^"]+)"/.exec(html), banco = bancoDeFotos(html);
  if(!mc || !banco) return null;
  var chave = mc[1], chaveC = chave + ":consulta", bancoC = banco + "-consulta";
  var h = html.split('"' + chave + '"').join('"' + chaveC + '"').split('"' + banco + '"').join('"' + bancoC + '"');
  if(h.indexOf('"' + chave + '"') >= 0 || h.indexOf('"' + banco + '"') >= 0) return null;   /* não isolou: não abre */
  var semente = { v: 1, lider: envio.lider || "", reg: envio.reg || {}, ocultos: {} };
  var inj = "<script>(function(){try{"
    + "var K=" + JSON.stringify(chaveC) + ";"
    + "Object.keys(localStorage).forEach(function(k){if(k.indexOf(K)===0)localStorage.removeItem(k);});"
    + "localStorage.setItem(K," + JSON.stringify(JSON.stringify(semente)).replace(/</g, "\\u003c") + ");"
    + "localStorage.setItem(K+':oc'," + JSON.stringify(JSON.stringify(envio.oc || [])).replace(/</g, "\\u003c") + ");"
    + "window.PLANTA_CONSULTA=true;"
    + "Object.defineProperty(navigator,'share',{configurable:true,value:function(){return Promise.reject(new DOMException('Só consulta','AbortError'));}});"
    + "Object.defineProperty(navigator,'canShare',{configurable:true,value:function(){return false;}});"
    + "document.addEventListener('click',function(e){var t=e.target&&e.target.closest&&e.target.closest('g.peca,#b-enviar,#faixa-fotos,#b-oc,#b-dados,#p-oc,#p-desfazer,#o-nova,#o-foto-bt,#env-abrir,#env-so-fotos,#env-fila-bt,[data-foto],.b-foto,input[type=file]');"
    + "if(t){e.stopPropagation();e.preventDefault();try{parent.Nucleo&&parent.Nucleo._avisoConsulta&&parent.Nucleo._avisoConsulta(t.matches&&t.matches('g.peca')?t.id:null);}catch(_){}}},true);"
    + "}catch(e){}})();<\/script>"
    + "<style>#b-enviar,#faixa-fotos,#faixa-par,#f-env,#b-oc,#b-dados,#p-oc,#p-desfazer,#o-nova,#o-foto-bt,.b-foto,input[type=file]{display:none!important}</style>";   /* #faixa-par: na consulta as fotos não ficam na planta, a conta dela sairia errada (ensaio 19/11) */
  var i = h.search(/<head[^>]*>/i), fim = i < 0 ? 0 : i + h.match(/<head[^>]*>/i)[0].length;
  return { html: h.slice(0, fim) + inj + h.slice(fim), chave: chaveC, banco: bancoC };
}
/* ------------------------------------------------ Sincronizar fotos (Gustavo, 18/11)
   As fotos vão DENTRO dos .txt da pasta Envios (base64). Cada envio leva só as fotos novas, então a
   obra inteira = as fotos de todos os envios dela. Sincronizar = baixar os envios desta planta que o
   aparelho ainda não tem e guardar as fotos num banco DA CASCA ("planta-online-fotos"), separado da
   planta: a consulta e o histórico mostram dali, sem marcar nem enviar nada. Ficam no aparelho até
   alguém apagar (serve sem sinal, para o histórico de "Minhas obras").
   Banco: fotos  {chave:"<obra>|<nome>", obra, nome, peca, alvo, quando, bytes, arquivo, blob}
          envios {chave:"<obra>|<arquivo>", obra, arquivo, tamanho, em, autor, fotos, lidoEm} */
var BANCO_FOTOS = "planta-online-fotos";
function abrirBancoFotos(){
  return new Promise(function(ok, falha){
    var r = indexedDB.open(BANCO_FOTOS, 1);
    r.onupgradeneeded = function(){ var db = r.result;
      if(!db.objectStoreNames.contains("fotos")) db.createObjectStore("fotos", { keyPath: "chave" }).createIndex("obra", "obra", { unique: false });
      if(!db.objectStoreNames.contains("envios")) db.createObjectStore("envios", { keyPath: "chave" }).createIndex("obra", "obra", { unique: false }); };
    r.onsuccess = function(){ ok(r.result); }; r.onerror = function(){ falha(r.error); };
  });
}
function todosDaObra(db, loja, obra){
  return new Promise(function(ok, falha){
    var q = db.transaction(loja).objectStore(loja).index("obra").getAll(obra);
    q.onsuccess = function(){ ok(q.result || []); }; q.onerror = function(){ falha(q.error); };
  });
}
function gravarNoBanco(db, fotos, envio){
  return new Promise(function(ok, falha){
    var tx = db.transaction(["fotos", "envios"], "readwrite"), st = tx.objectStore("fotos");
    var c = st.index("obra").openCursor(IDBKeyRange.only(envio.obra));
    c.onsuccess = function(e){ var cur = e.target.result;
      if(cur){ if(cur.value.arquivo === envio.arquivo) cur.delete(); cur.continue(); return; }
      fotos.forEach(function(f){ st.put(f); }); };
    tx.objectStore("envios").put(envio);
    tx.oncomplete = function(){ ok(); }; tx.onerror = function(){ falha(tx.error); };
  });
}
function base64ParaBlob(s){
  s = String(s || ""); if(s.indexOf("data:") === 0) s = s.slice(s.indexOf(",") + 1);
  var bin = atob(s), u = new Uint8Array(bin.length);
  for(var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
  return new Blob([u], { type: "image/jpeg" });
}
/* o que falta baixar: envios .txt desta planta que o aparelho ainda não leu (ou que mudaram de tamanho)
   deps = {listar, baixarInicio, identidade} -> Promise<{novos:[item], bytes, jaTem, total}> */
function planoDeSincronia(o, deps){
  return Promise.all([deps.listar(o.envios), abrirBancoFotos().then(function(db){ return todosDaObra(db, "envios", o.id).then(function(l){ db.close(); return l; }); })])
    .then(function(x){
      var itens = (x[0] || []).filter(function(i){ return /\.(txt|json)$/i.test(i.name); }), lidos = {};   /* mesmo critério do envios_planta.py */
      x[1].forEach(function(e){ lidos[e.arquivo] = e; });
      var faltam = itens.filter(function(i){ var e = lidos[i.name]; return !e || e.tamanho !== i.size || (i.eTag && e.eTag && e.eTag !== i.eTag); });
      /* pasta dividida entre obras: confere a identidade no começo do arquivo antes de baixar tudo */
      return faltam.reduce(function(ac, i){
        return ac.then(function(l){
          if(!deps.identidade) { l.push(i); return l; }
          return deps.baixarInicio(i.url, 4096).then(function(t){
            var c = lerCabeca(t);
            if(FORMATOS_COM_ADAPTADOR.indexOf(c.formato) >= 0 && mesmaPlanta(c, deps.identidade)) l.push(i);
            return l;
          }, function(){ return l; });
        });
      }, Promise.resolve([])).then(function(novos){
        var bytes = 0; novos.forEach(function(i){ bytes += i.size || 0; });
        return { novos: novos, bytes: bytes, jaTem: x[1].length, total: itens.length };
      });
    });
}
/* baixa os envios do plano e guarda as fotos. deps.baixarTudo(url) -> Promise<texto>; aoProgresso(i, n, item) */
/* deps.cancelado() -> true quando a consulta foi fechada: nada mais é gravado (achado B5b da validação) */
function sincronizarFotos(o, plano, deps, aoProgresso){
  var res = { envios: 0, fotos: 0, falhas: [], cancelado: false };
  function parou(){ if(deps.cancelado && deps.cancelado()){ res.cancelado = true; return true; } return false; }
  return abrirBancoFotos().then(function(db){
    return plano.novos.reduce(function(ac, item, k){
      return ac.then(function(){
        if(parou()) return;
        if(aoProgresso) aoProgresso(k, plano.novos.length, item);
        return deps.baixarTudo(item.url).then(function(t){
          if(parou()) return;
          var d = JSON.parse(t), autor = (item.createdBy && item.createdBy.user) || {};
          var fotos = (d.fotos || []).filter(function(f){ return f && f.dados; }).map(function(f){
            return { chave: o.id + "|" + (f.nome || item.name + "#" + Math.random()), obra: o.id, nome: f.nome || null,
                     peca: f.peca || null, alvo: f.alvo || null, quando: f.quando || null, bytes: f.bytes || null,
                     arquivo: item.name, blob: base64ParaBlob(f.dados) };
          });
          return gravarNoBanco(db, fotos, { chave: o.id + "|" + item.name, obra: o.id, arquivo: item.name, tamanho: item.size, eTag: item.eTag || null,
                                            em: item.createdDateTime || null, autor: autor.displayName || null,
                                            fotos: fotos.length, lidoEm: new Date().toISOString() })
            .then(function(){ res.envios++; res.fotos += fotos.length; });
        }).catch(function(e){ res.falhas.push({ arquivo: item.name, erro: String(e && e.message || e),
                                                tipo: e instanceof SyntaxError ? "arquivo_ruim" : "rede" }); });
      });
    }, Promise.resolve()).then(function(){ db.close(); return res; });
  });
}
/* Obra no histórico: as fotos NÃO ficam no aparelho (Gustavo, 18/11). Quem quiser ver sincroniza de novo.
   Apaga só a cópia da casca (que veio do SharePoint); a planta do líder não é tocada. */
function apagarFotosSincronizadas(obraIds){
  obraIds = [].concat(obraIds || []).filter(Boolean);
  if(!obraIds.length) return Promise.resolve(0);
  return (indexedDB.databases ? indexedDB.databases() : Promise.resolve(null)).then(function(ds){
    if(ds && !ds.some(function(d){ return d.name === BANCO_FOTOS; })) return 0;
    return abrirBancoFotos().then(function(db){
      return new Promise(function(ok){
        var n = 0, tx = db.transaction(["fotos", "envios"], "readwrite");
        ["fotos", "envios"].forEach(function(loja){
          obraIds.forEach(function(id){
            var c = tx.objectStore(loja).index("obra").openCursor(IDBKeyRange.only(id));
            c.onsuccess = function(e){ var cur = e.target.result; if(!cur) return; if(loja === "fotos") n++; cur.delete(); cur.continue(); };
          });
        });
        tx.oncomplete = function(){ db.close(); ok(n); }; tx.onerror = function(){ db.close(); ok(n); };
      });
    });
  }).catch(function(){ return 0; });
}
/* -> Promise<{fotos:[...com blob], envios:[...], ultimoEnvio, autor, sincronizadoEm}> */
function fotosSincronizadas(obraId){
  return abrirBancoFotos().then(function(db){
    return Promise.all([todosDaObra(db, "fotos", obraId), todosDaObra(db, "envios", obraId)]).then(function(x){
      db.close();
      var env = x[1].slice().sort(function(a, b){ return String(a.em).localeCompare(String(b.em)); }), ult = env[env.length - 1] || null;
      var fotos = x[0].slice().sort(function(a, b){ return String(a.peca).localeCompare(String(b.peca)) || String(a.quando).localeCompare(String(b.quando)); });
      var quando = null; env.forEach(function(e){ if(!quando || e.lidoEm > quando) quando = e.lidoEm; });
      return { fotos: fotos, envios: env, ultimoEnvio: ult ? ult.em : null, autor: ult ? ult.autor : null, sincronizadoEm: quando };
    });
  }).catch(function(){ return { fotos: [], envios: [], ultimoEnvio: null, autor: null, sincronizadoEm: null }; });
}

/* Planta de DEMONSTRAÇÃO ("Ver como" do ADM, só no teste): a planta inteira, tocável, mas isolada (CHAVE e banco
   com sufixo ":simulacao"/"-simulacao") e sem envio (o compartilhar recusa). Sai com limparConsulta. */
function htmlDeSimulacao(html){
  if(FORMATOS_COM_ADAPTADOR.indexOf(formatoDaPlanta(html)) < 0) return null;
  var mc = /var CHAVE\s*=\s*"([^"]+)"/.exec(html), banco = bancoDeFotos(html);
  if(!mc || !banco) return null;
  var chave = mc[1], chaveS = chave + ":simulacao", bancoS = banco + "-simulacao";
  var h = html.split('"' + chave + '"').join('"' + chaveS + '"').split('"' + banco + '"').join('"' + bancoS + '"');
  if(h.indexOf('"' + chave + '"') >= 0 || h.indexOf('"' + banco + '"') >= 0) return null;
  var inj = "<script>(function(){try{window.PLANTA_SIMULACAO=true;"
    + "var nao=function(){try{parent.Nucleo&&parent.Nucleo._avisoSimulacao&&parent.Nucleo._avisoSimulacao();}catch(_){}return Promise.reject(new DOMException('Demonstração','AbortError'));};"
    + "Object.defineProperty(navigator,'share',{configurable:true,value:nao});"
    + "Object.defineProperty(navigator,'canShare',{configurable:true,value:function(){return true;}});"
    + "}catch(e){}})();<\/script>";
  var i = h.search(/<head[^>]*>/i), fim = i < 0 ? 0 : i + h.match(/<head[^>]*>/i)[0].length;
  return { html: h.slice(0, fim) + inj + h.slice(fim), chave: chaveS, banco: bancoS };
}
/* ------------------------------------------------ arquivamento da obra no aparelho (Gustavo, 19/11)
   4 semanas depois do término (encerrada_em), a planta sai do aparelho do líder. A partir de 7 dias antes
   ele é avisado e pode ADIAR (+14 dias, quantas vezes quiser) ou ARQUIVAR AGORA.
   Só apaga o que está PROVADO no SharePoint (decisão de 18/11): cada foto do aparelho tem que estar na
   lista "fotosArquivos" de algum envio da obra, e nenhuma marcação pode estar pendente. Sem prova, não apaga.
   Apaga: o que a planta guarda (localStorage CHAVE*, banco de fotos), a planta guardada (cache), a cópia de
   fotos sincronizadas e os registros da casca da obra. */
var DIAS_ARQUIVO = 28, DIAS_AVISO = 7, DIAS_ADIAR = 14;
/* DESLIGADO na 0.3 (revisão do ChatGPT, 19/11, X1/X2): a prova por nome em "fotosArquivos" não garante que a foto
   é recuperável, nem cobre marcação local sem envio. Só religar com prova dos BYTES de cada foto (e do estado
   remoto igual ao local) e teste em aparelho real. Enquanto false: nada é apagado, nem botão aparece. */
var ARQUIVAMENTO_ATIVO = false;
function somarDias(iso, d){ var t = new Date(iso + (String(iso).length === 10 ? "T12:00:00" : "")); t.setDate(t.getDate() + d); return t.toISOString().slice(0, 10); }
function prazoArquivo(o){
  if(!o || !o.encerrada_em) return null;
  var adiado = lerLS("arquivo:" + o.id, null), base = somarDias(o.encerrada_em, DIAS_ARQUIVO);
  return adiado && adiado.ate && adiado.ate > base ? adiado.ate : base;
}
function adiarArquivo(o){
  var atual = prazoArquivo(o) || hojeISO(), novo = somarDias(atual > hojeISO() ? atual : hojeISO(), DIAS_ADIAR);
  gravarLS("arquivo:" + o.id, { ate: novo, em: new Date().toISOString() }); return novo;
}
function diasAte(iso){ return Math.round((Date.parse(iso + "T12:00:00") - Date.parse(hojeISO() + "T12:00:00")) / 86400000); }
/* a obra tem algo neste aparelho? (info guardada ao abrir, registro de envio ou planta no cache) */
function temNoAparelho(id){ return !!(lerLS("info:" + id, null) || lerLS("envio:" + id, null)); }
/* fotos que a planta guarda no aparelho: [{id, nome?, peca, alvo, quando, enviada}] (nome é o que a planta dá) */
function fotosDoBanco(banco){
  return new Promise(function(ok){
    var lista = indexedDB.databases ? indexedDB.databases() : Promise.resolve(null);
    lista.then(function(ds){
      if(ds && !ds.some(function(d){ return d.name === banco; })) return ok([]);
      var r = indexedDB.open(banco);
      r.onsuccess = function(){ var db = r.result;
        if(!db.objectStoreNames.contains("fotos")){ db.close(); return ok([]); }
        var q = db.transaction("fotos").objectStore("fotos").getAll();
        q.onsuccess = function(){ db.close(); ok((q.result || []).map(function(f){ return { id: f.id, nome: f.nome || null, peca: f.peca, alvo: f.alvo, quando: f.quando, enviada: !!f.enviada, rol: f.rol || null }; })); };
        q.onerror = function(){ db.close(); ok(null); }; };
      r.onerror = function(){ ok(null); };
    }, function(){ ok(null); });
  });
}
/* nomes de foto que chegaram, pelos envios da obra (lista fotosArquivos do começo de cada pacote) */
function nomesEnviados(o, deps){
  return deps.listar(o.envios).then(function(itens){
    itens = (itens || []).filter(function(x){ return /\.(txt|json)$/i.test(x.name); });
    var nomes = {};
    return itens.reduce(function(ac, x){
      return ac.then(function(){
        return lerInicio(x, deps).then(function(l){
          var c = l && l.c; if(!c) return;
          if(deps.identidade && !mesmaPlanta(c, deps.identidade)) return;
          var fa = valorJSON(l.texto, "fotosArquivos");
          if(fa === undefined && (x.size || 0) > CABECA_PEQUENA) return deps.baixarInicio(x.url, CABECA).then(function(t2){ fa = valorJSON(t2, "fotosArquivos"); (fa || []).forEach(function(f){ nomes[String(f.nome || "").replace(/\.jpe?g$/i, "")] = true; }); });
          (fa || []).forEach(function(f){ nomes[String(f.nome || "").replace(/\.jpe?g$/i, "")] = true; });
        }, function(){});
      });
    }, Promise.resolve()).then(function(){ return nomes; });
  });
}
/* -> Promise<{ok, motivo, fotosNoAparelho, semProva:[...], pendencia}> */
function provarArquivamento(o, deps){
  if(!ARQUIVAMENTO_ATIVO) return Promise.resolve({ ok: false, motivo: "arquivamento desligado nesta versão", fotosNoAparelho: null, semProva: [], pendencia: null });
  var info = lerLS("info:" + o.id, null);
  if(!info) return Promise.resolve({ ok: true, motivo: "nada da planta neste aparelho", fotosNoAparelho: 0, semProva: [], pendencia: false });
  return Promise.all([pendenciaDoAparelho(o.id), fotosDoBanco(info.banco)]).then(function(x){
    var pend = x[0], fotos = x[1];
    if(fotos === null) return { ok: false, motivo: "não deu para ler as fotos do aparelho", fotosNoAparelho: null, semProva: [], pendencia: pend };
    if(pend) return { ok: false, motivo: "há marcação ou foto que não foi enviada", fotosNoAparelho: fotos.length, semProva: [], pendencia: true };
    if(!fotos.length) return { ok: true, motivo: "sem fotos no aparelho", fotosNoAparelho: 0, semProva: [], pendencia: false };
    return nomesEnviados(o, deps).then(function(nomes){
      var sem = fotos.filter(function(f){ return !f.enviada || !f.nome || !nomes[String(f.nome).replace(/\.jpe?g$/i, "")]; });
      return { ok: !sem.length, motivo: sem.length ? sem.length + " foto(s) sem cópia comprovada no SharePoint" : "todas as fotos estão nos envios",
               fotosNoAparelho: fotos.length, semProva: sem.map(function(f){ return f.nome || ("foto " + f.id); }), pendencia: false };
    });
  });
}
/* apaga do aparelho (só depois de provarArquivamento ok) */
function arquivarObra(o){
  if(!ARQUIVAMENTO_ATIVO) return Promise.reject(new Error("arquivamento desligado nesta versão"));
  var info = lerLS("info:" + o.id, null), feitos = [];
  try{
    if(info && info.chave){ Object.keys(localStorage).forEach(function(k){ if(k === info.chave || k.indexOf(info.chave + ":") === 0){ localStorage.removeItem(k); } }); feitos.push("marcações"); }
    ["envio:", "info:", "arquivo:"].forEach(function(p){ localStorage.removeItem("online:" + p + o.id); });
  }catch(e){}
  var lista = lerLS("arquivadas", []); if(lista.indexOf(o.id) < 0){ lista.push(o.id); gravarLS("arquivadas", lista); }
  var apagaBanco = info && info.banco ? new Promise(function(ok){ try{ var r = indexedDB.deleteDatabase(info.banco); r.onsuccess = r.onerror = r.onblocked = function(){ ok(); }; }catch(e){ ok(); } }) : Promise.resolve();
  var apagaCache = (window.caches ? caches.open("plantas").then(function(c){ return c.delete("planta/" + o.id); }) : Promise.resolve()).catch(function(){});
  return Promise.all([apagaBanco, apagaCache, apagarFotosSincronizadas(o.id)]).then(function(){ feitos.push("fotos", "planta guardada"); return feitos; });
}
/* obras do histórico com algo neste aparelho, com o prazo e quantos dias faltam */
function obrasParaArquivar(manifesto){
  if(!ARQUIVAMENTO_ATIVO) return [];
  return ((manifesto && manifesto.encerradas) || []).filter(function(o){ return temNoAparelho(o.id); })
    .map(function(o){ var p = prazoArquivo(o); return { obra: o, prazo: p, dias: p ? diasAte(p) : null }; });
}
/* apaga o que a consulta deixou no aparelho */
function limparConsulta(c){
  if(!c) return Promise.resolve();
  try{ Object.keys(localStorage).forEach(function(k){ if(k.indexOf(c.chave) === 0) localStorage.removeItem(k); }); }catch(e){}
  return new Promise(function(ok){ try{ var r = indexedDB.deleteDatabase(c.banco); r.onsuccess = r.onerror = r.onblocked = function(){ ok(); }; }catch(e){ ok(); } });
}

window.Nucleo = {
  ultimoEnvioConfirmado: ultimoEnvioConfirmado, htmlDeConsulta: htmlDeConsulta, limparConsulta: limparConsulta,
  _lerCabeca: lerCabeca, _valorJSON: valorJSON,
  versao: "0.3-lab",
  carregarAcessos: carregarAcessos, carregarEditores: carregarEditores,
  perfilDaConta: perfilDaConta, listasDaConta: listasDaConta, editorDaObra: editorDaObra,
  identidadeDaPlanta: identidadeDaPlanta, arquivamentoAtivo: function(){ return ARQUIVAMENTO_ATIVO; }, prazoArquivo: prazoArquivo, adiarArquivo: adiarArquivo, provarArquivamento: provarArquivamento, arquivarObra: arquivarObra, obrasParaArquivar: obrasParaArquivar, temNoAparelho: temNoAparelho, htmlDeSimulacao: htmlDeSimulacao, carregarAjustes: carregarAjustes, novoAjuste: novoAjuste, pastaHistorico: pastaHistorico, historicoDosLideres: historicoDosLideres, pendenciaDoAparelho: pendenciaDoAparelho, lembrarPlanta: lembrarPlanta, temPendencia: temPendencia, parteDoNome: parteDoNome, apagarFotosSincronizadas: apagarFotosSincronizadas, nomeDoLiderNaPlanta: nomeDoLiderNaPlanta, planoDeSincronia: planoDeSincronia, sincronizarFotos: sincronizarFotos, fotosSincronizadas: fotosSincronizadas,   carregarTrocas: carregarTrocas, pastaTrocas: pastaTrocas, novaTroca: novaTroca, podeAssumir: podeAssumir, podeMarcar: podeMarcar, resolverNome: resolverNome,
  modoDaObra: modoDaObra, formatoDaPlanta: formatoDaPlanta,
  envioComecou: envioComecou, envioTerminou: envioTerminou, observar: observar, estadoEnvio: estadoEnvio,
  infoDaPlanta: infoDaPlanta, lerEstadoPlanta: lerEstadoPlanta,
  cabecalhoDoPacote: cabecalhoDoPacote,
  espacoUsado: espacoUsado, bancoDeFotos: bancoDeFotos,
  _norm: norm, _mesmoNome: mesmoNome
};
})();
