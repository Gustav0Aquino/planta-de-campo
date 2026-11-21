/* Planta de Campo online: a casca (30/10/2026).
   A planta continua sendo o HTML gerado de cada obra, sem mudança. A casca:
   - faz o login Microsoft (app "Planta de Campo (protótipo)", Sites.Selected no site Montagem);
   - lê a lista de obras (obras.json no SharePoint) e abre a planta da obra num iframe;
   - guarda a planta para abrir sem sinal;
   - troca o compartilhar do Android (navigator.share) por gravação direta na Envios da obra.
     O "share" só termina depois que o SharePoint confirma o arquivo inteiro: a planta só marca
     "enviado" quando chegou de verdade. Sem sinal, espera o sinal voltar com a planta aberta.
   Os arquivos são os mesmos que o líder mandava pelo compartilhar: os importadores não mudam. */
var VERSAO = "0.3";
var CLIENTE = "be9bf767-d1a4-4735-adec-cf4aacc40391";
var LOCATARIO = "8999042a-e592-47e6-a75e-58f8901c3645";
var SITE_ID = "montaggiotech.sharepoint.com,b0667df5-71a6-4efd-a8b1-b02d9192ccb1,4d2f1e7c-b7be-4cf5-b3b0-12cb70c6febd";
var ESCOPOS = ["User.Read", "Sites.Selected"];
/* lista de obras: produção e teste (abrir com #teste liga o teste neste aparelho; #producao desliga) */
var MANIFESTOS = { producao: "General/_PLANTA ONLINE/obras.json", teste: "General/_TESTE PLANTA/obras.json" };
var PEDACO = 10 * 327680; /* a sessão de envio do Graph pede pedaços múltiplos de 320 KiB */

var $ = function(id){ return document.getElementById(id); };
if(!window.Nucleo){
  ["t-entrar", "t-obras"].forEach(function(id){ var e = document.getElementById(id); if(e) e.hidden = true; });
  var tm = document.getElementById("t-msg"); if(tm){ tm.hidden = false;
    document.getElementById("msg-tit").textContent = "Não deu para abrir o app";
    document.getElementById("msg-txt").textContent = "Feche e abra de novo com sinal.";
    var bt = document.getElementById("msg-bt"); bt.hidden = false; bt.textContent = "Tentar de novo"; bt.onclick = function(){ location.reload(); }; }
  throw new Error("nucleo.js não carregou");
}
function ler(k, padrao){ try{ var v = localStorage.getItem("online:" + k); return v == null ? padrao : JSON.parse(v); }catch(e){ return padrao; } }
function gravar(k, v){ try{ localStorage.setItem("online:" + k, JSON.stringify(v)); }catch(e){} }
function registro(t){ var l = ler("registro", []); l.push(new Date().toISOString() + " " + t); if(l.length > 80) l = l.slice(-80); gravar("registro", l); }
/* 0.2 (C19): texto que vem de fora (erro da Microsoft, nome de arquivo) entra no aviso como texto, nunca como HTML */
function esc(s){ return String(s == null ? "" : s).replace(/[&<>"']/g, function(c){ return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
function hora(d){ return new Date(d).toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" }); }
function mb(b){ return (b / 1048576).toFixed(1).replace(".", ","); }
function espera(ms){ return new Promise(function(ok){ setTimeout(ok, ms); }); }
function tela(id){
  ["t-entrar", "t-obras", "t-msg"].forEach(function(t){ $(t).hidden = t !== id; });
  if(id !== "t-obras") fecharPerfil();
  $("planta").hidden = $("faixa-obra").hidden = $("campo").hidden = !!id;
}
function mensagem(tit, txt, tentar, acao){
  tela("t-msg"); $("msg-tit").textContent = tit; $("msg-txt").innerHTML = txt || "";
  $("msg-bt").textContent = acao || "Tentar de novo";
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
  /* 0.2 (05/11, laboratório B8d): uma lista sem obras (arquivo salvo pela metade, "{}", modelo errado) NÃO
     substitui a lista boa guardada; sem isso o líder perdia a obra até sem sinal. BOM do Bloco de Notas é ignorado. */
  return baixarTexto(MANIFESTOS[MODO]).then(function(r){
    var m = JSON.parse(String(r.texto).replace(/^﻿/, ""));
    var ok = m && Array.isArray(m.obras) && m.obras.length && m.obras.every(function(o){ return o && o.id && o.planta && o.envios; });
    if(!ok) throw new Error("lista de obras sem obras ou com obra incompleta (id, planta, envios)");
    gravar("obras:" + MODO, m); return m; })
    .catch(function(e){ var m = ler("obras:" + MODO, null); registro("lista de obras: " + e.message); if(m) return m; throw e; });
}
/* 0.3 (18/11): perfil do líder. Minhas obras / Obras de outros líderes / Histórico.
   As decisões vêm do Nucleo (nucleo.js); este desenho de tela é PROVISÓRIO: o design é do Codex. */
var EXTRAS = { acessos: null, editores: null, trocas: [], ajustes: [], origemAcessos: null };
function baixarUrl(url){ return fetch(url).then(function(r){ if(!r.ok) throw new Error("download " + r.status); return r.text(); }); }
function carregarExtras(){
  return Promise.all([Nucleo.carregarAcessos(baixarTexto, MANIFESTOS[MODO]), Nucleo.carregarEditores(baixarTexto, MANIFESTOS[MODO])])
    .then(function(x){ EXTRAS.acessos = x[0].dados; EXTRAS.origemAcessos = x[0].origem; EXTRAS.editores = x[1].dados;
      return Promise.all([Nucleo.carregarTrocas(listarPasta, MANIFESTOS[MODO]), Nucleo.carregarAjustes(listarPasta, baixarUrl, MANIFESTOS[MODO])]).then(function(y){
        var tr = y[0]; EXTRAS.trocas = tr.dados || []; EXTRAS.ajustes = y[1].dados || [];
        registro("extras: acessos " + x[0].origem + ", editores " + x[1].origem + ", trocas " + tr.origem + " (" + EXTRAS.trocas.length + "), histórico " + y[1].origem + " (" + EXTRAS.ajustes.length + ")"); return EXTRAS; }); })
    .catch(function(e){ registro("extras: " + (e.message || e)); if(!EXTRAS.acessos) EXTRAS.origemAcessos = "erro"; return EXTRAS; });
}
/* "Ver como" (Gustavo, 19/11): só no modo TESTE e só para ADM. A tela mostra o que aquele líder vê, mas nada é
   gravado em nome dele: abrir vira consulta e a troca fica desligada. */
var VISAO = null;
function perfilReal(){ return Nucleo.perfilDaConta(conta(), EXTRAS.acessos, EXTRAS.origemAcessos); }
function perfil(){
  var r = perfilReal();
  if(!VISAO || !r.adm || MODO !== "teste") return r;
  var v = Nucleo.perfilDaConta({ name: VISAO.nome, username: VISAO.email }, EXTRAS.acessos, EXTRAS.origemAcessos);
  v.simulado = true; return v;
}
function ctxObras(){ return { editores: EXTRAS.editores, trocas: EXTRAS.trocas, acessos: EXTRAS.acessos, ajustes: EXTRAS.ajustes }; }
function cartaoObra(tit, sub, acao, fn, cls){
  var b = document.createElement(fn ? "button" : "article");
  b.className = ["obra", cls, fn ? "" : "estatica"].filter(Boolean).join(" ");
  var nome = document.createElement("span"), detalhe = document.createElement("small");
  nome.className = "obra-nome";
  nome.textContent = tit;
  detalhe.textContent = sub;
  b.append(nome, detalhe);
  if(fn){
    var rotulo = document.createElement("span");
    rotulo.className = "obra-acao";
    rotulo.textContent = acao;
    b.appendChild(rotulo);
    b.type = "button";
    b.onclick = fn;
  }
  return b;
}
function tituloGrupo(txt){ var h = document.createElement("h2"); h.className = "grupo"; h.textContent = txt; return h; }
function vazio(txt){ var p = document.createElement("p"); p.className = "vazio"; p.textContent = txt; return p; }
function dataHistorico(iso){ return iso.replace(/^(\d{4})-(\d{2})-(\d{2})$/, "$3/$2/$1"); }
function mostrarObras(m){
  /* 18/11 (Gustavo, no teste): a lista de obras volta a ser a da 0.2 (todas as obras que o Gustavo sobe,
     no azul da Montaggio). O perfil do líder fica numa aba ao lado: "Minhas obras", com as obras em
     andamento dele e as que já terminaram. Obra de outro líder continua abrindo em "Só consulta". */
  tela("t-obras");
  var p = perfil(), real = perfilReal(), L = $("lista-obras"), P = $("lista-perfil");
  $("quem").textContent = p.simulado ? "Conta de " + real.nome + " · vendo como " + p.nome : (p.nome ? "Conta de " + p.nome : "");
  L.replaceChildren(); P.replaceChildren();
  if(real.adm && MODO === "teste" && EXTRAS.acessos){
    var caixa = document.createElement("label"); caixa.className = "ver-como";
    caixa.append("Ver como: ");
    var sel = document.createElement("select"); sel.id = "ver-como";
    var o0 = document.createElement("option"); o0.value = ""; o0.textContent = "Minha visão (ADM)"; sel.appendChild(o0);
    (EXTRAS.acessos.lideres || []).forEach(function(l){ var op = document.createElement("option"); op.value = l.email; op.textContent = l.nome;
      if(VISAO && VISAO.email === l.email) op.selected = true; sel.appendChild(op); });
    sel.onchange = function(){ var l = (EXTRAS.acessos.lideres || []).filter(function(x){ return x.email === sel.value; })[0];
      VISAO = l ? { nome: l.nome, email: l.email } : null; mostrarObras(m); };
    caixa.appendChild(sel); L.appendChild(caixa);
    if(p.simulado) L.appendChild(vazio("Visão simulada de " + p.nome + " (só no teste). Nada é gravado: \"Abrir\" abre uma demonstração da planta (marca e tira foto, mas não envia) e a troca de líder fica desligada."));
  }
  /* na visão simulada, a obra "somenteConsulta" da lista de teste aparece como o líder a vê na produção
     (é seguro: na simulação, "Abrir" é só a demonstração, que não grava nem envia) */
  var mVis = p.simulado ? Object.assign({}, m, { obras: (m.obras || []).map(function(o){ var c = Object.assign({}, o); delete c.somenteConsulta; return c; }) }) : m;
  var ls = Nucleo.listasDaConta(mVis, p, ctxObras());
  Nucleo.apagarFotosSincronizadas((m.encerradas || []).map(function(x){ return x.id; }));
  function resp(it){ return !it.editor ? "Responsável não definido"
    : it.editor.naoIdentificado ? "Responsável não identificado (" + it.editor.nome + ") · avise o Gustavo"
    : "Responsável: " + it.editor.nome + (it.confirmado ? "" : " (pelo nome)"); }
  if(p.papel === "indeterminado") L.appendChild(vazio("Não deu para confirmar o seu perfil (sem sinal?). Por enquanto, só consulta. Abra o app de novo com sinal."));
  /* obra que não é (mais) desta conta, mas com marcação ou foto deste aparelho que não subiu: sempre há como enviar */
  function botaoPendencia(o, ed, depoisDe, lista){
    if(!o.planta || !o.envios || p.simulado) return;
    Nucleo.pendenciaDoAparelho(o.id).then(function(tem){
      if(!tem || !document.body.contains(depoisDe)) return;
      var x = cartaoObra(o.nome, "Você tem marcações ou fotos não enviadas desta obra neste aparelho · abra só para ENVIAR",
        "Abrir", function(){ fecharPerfil(); abrirObra(o, { pendencias: true, lider: ed && ed.nome }); }, "pend");
      (lista || L).insertBefore(x, depoisDe.nextSibling);
    });
  }
  function marcar(o){ fecharPerfil(); if(p.simulado){ abrirSimulacao(o, p); return; } gravar("ultima", o.id); abrirObra(o); }
  /* lista principal: as obras na ordem do portal; as do líder abrem para marcar, as dos outros em consulta */
  var minhas = {}; ls.minhas.forEach(function(it){ minhas[it.obra.id] = it; });
  var outras = {}; ls.outras.forEach(function(it){ outras[it.obra.id] = it; });
  (m.obras || []).forEach(function(o){
    var atual = o.id === ler("ultima", null);
    if(minhas[o.id]){
      L.appendChild(cartaoObra(o.nome, [resp(minhas[o.id]), o.periodo, atual ? "Aberta por último" : ""].filter(Boolean).join(" · "),
        "Abrir", function(){ marcar(o); }, atual ? "atual" : ""));
      return;
    }
    var it = outras[o.id]; if(!it) return;
    var cartao = cartaoObra(o.nome, [resp(it), o.periodo].filter(Boolean).join(" · "),
      "Só consulta", function(){ abrirConsulta(o, it); }, "outra");
    L.appendChild(cartao);
    botaoPendencia(o, it.editor, cartao, L);
    if(it.podeAssumir){
      var x = document.createElement("button");
      x.className = "link"; x.type = "button";
      x.textContent = "Agora sou o líder desta obra";
      x.setAttribute("aria-label", x.textContent + " · " + o.nome);
      x.onclick = function(){ if(p.simulado){ aviso("Na visão simulada a troca de líder fica desligada.", ""); avisoSome(3500); return; } assumirObra(o, it); };
      L.appendChild(x);
    }
  });
  if(!L.children.length) L.appendChild(vazio("Nenhuma obra na lista."));
  if(!p.simulado) avisosDeArquivo(m, L);
  /* obra que foi para o histórico com algo não enviado neste aparelho: aparece na lista principal só para enviar */
  (m.encerradas || []).forEach(function(o){ var marca = document.createElement("span"); L.appendChild(marca); botaoPendencia(o, null, marca, L); });
  /* aba "Minhas obras" */
  P.appendChild(tituloGrupo("Em andamento"));
  if(!ls.minhas.length) P.appendChild(vazio("Nenhuma obra em andamento no seu nome."));
  ls.minhas.forEach(function(it){
    var o = it.obra;
    P.appendChild(cartaoObra(o.nome, [resp(it), o.periodo].filter(Boolean).join(" · "), "Abrir", function(){ marcar(o); }, ""));
  });
  P.appendChild(tituloGrupo("Terminadas"));
  if(!ls.realizadas.length) P.appendChild(vazio("Nenhuma obra terminada ainda."));
  ls.realizadas.forEach(function(h){
    var o = h.obra, cod = o.codigo && !o.nome.includes(o.codigo) ? o.codigo + " · " : "";
    var per = h.de ? dataHistorico(h.de) + " a " + (h.ate ? dataHistorico(h.ate) : "hoje")
      : h.ate ? "Até " + dataHistorico(h.ate) : "Período não informado";
    P.appendChild(cartaoObra(cod + o.nome,
      (h.papel === "adm" ? ["Líder: " + ((h.lideres || []).join(", ") || "não informado"), h.ate ? "Encerrada em " + dataHistorico(h.ate) : "Encerrada", "fotos: sincronizar para ver"]
        : [h.papel === "escalado" ? "Escalado" : "Editou", per, h.ativa ? "Obra ainda em andamento" : "Encerrada", h.confirmado ? "" : "Pelo nome"]).filter(Boolean).join(" · "),
      "Só consulta", (h.ativa || (o.planta && o.envios)) ? function(){ fecharPerfil(); abrirConsulta(o); } : null, "hist"));
    if(!h.ativa && !p.simulado && Nucleo.arquivamentoAtivo() && Nucleo.temNoAparelho(o.id)){
      var lib = document.createElement("button"); lib.type = "button"; lib.className = "link";
      lib.textContent = "Liberar espaço: tirar esta planta do aparelho";
      lib.onclick = function(){ arquivarAgora(m, o); }; P.appendChild(lib);
    }

  });
  if(real.adm && !p.simulado && EXTRAS.acessos) painelLideres(m, P);
  $("b-perfil").querySelector("b").textContent = ls.minhas.length ? String(ls.minhas.length) : "";
  window.scrollTo(0, 0);
}
/* "Agora sou o líder" (Gustavo, 18/11): confirmação em dois passos, com o botão de confirmar em lugares
   diferentes, para um toque acidental (ou dois seguidos) não trocar o líder. Grava um arquivo novo em trocas/
   (nunca sobrescreve); quem trocou fica no "Criado por" do SharePoint. */
function confirmar(tit, txt, sim, nao, lado){
  return new Promise(function(ok){
    var c = $("confirma"); $("conf-tit").textContent = tit; $("conf-txt").textContent = txt;
    var bs = $("conf-bts"); bs.replaceChildren(); bs.className = lado || "";
    function bt(rot, cls, v){ var b = document.createElement("button"); b.type = "button"; b.className = cls; b.textContent = rot;
      b.onclick = function(){ c.hidden = true; ok(v); }; return b; }
    var bSim = bt(sim, "conf-sim", true), bNao = bt(nao, "conf-nao", false);
    bSim.disabled = true; setTimeout(function(){ bSim.disabled = false; }, 900);   /* segundo toque rápido não confirma */
    if(lado === "inverte") bs.append(bSim, bNao); else bs.append(bNao, bSim);
    c.hidden = false;
  });
}
function assumirObra(o, it){
  var p = perfil(), ant = it && it.editor, quem = ant && ant.nome ? ant.nome : "ninguém";
  confirmar("Trocar o líder desta obra?",
    o.nome + "\n\nHoje o líder é: " + quem + ".\nSe você continuar, você passa a marcar e enviar esta obra. " + quem + " passa a ver a obra em consulta e ainda pode enviar o que já tinha marcado.",
    "Continuar", "Cancelar").then(function(v1){
    if(!v1) return;
    return confirmar("Confirme de novo",
      "Você, " + (p.nome || p.email) + ", é o líder de " + o.nome + " a partir de agora?",
      "Sim, sou o líder", "Não", "inverte").then(function(v2){
      if(!v2) return;
      var t = Nucleo.novaTroca(o, p, ant), f = new File([t.conteudo], t.nome, { type: "application/json" });
      mensagem("Trocando o líder…", "");
      function souEu(){ var ed = Nucleo.editorDaObra(o, ctxObras()); return ed && ed.email && ed.email.toLowerCase() === String(p.email).toLowerCase() ? null : (ed || {}); }
      function voltarLista(txt){
        aviso(txt, "erro", [{ txt: "OK", fn: function(){ aviso(null); } }]);
        return carregarObras().then(mostrarObras, function(){ var m = ler("obras:" + MODO, null); if(m) mostrarObras(m); });
      }
      /* antes de gravar: alguém assumiu enquanto a tela estava aberta? */
      return carregarExtras().then(function(){
        var agora = Nucleo.editorDaObra(o, ctxObras());
        if(ant && agora && (agora.email || agora.nome) !== (ant.email || ant.nome))
          return voltarLista("<b>" + esc(agora.nome || agora.email) + " assumiu esta obra agora.</b> Nada foi trocado. Confira a lista e, se precisar, tente de novo.");
        return enviarArquivo(f, Nucleo.pastaTrocas(MANIFESTOS[MODO]) + "/" + t.nome, { agora: false, silencioso: true }).then(function(){
          registro("troca de líder: " + o.id + " → " + p.email);
          return carregarExtras().then(function(){
            var outro = souEu();
            if(outro) return voltarLista("<b>" + esc(outro.nome || outro.email || "Outra pessoa") + " assumiu esta obra ao mesmo tempo, e ficou com ela.</b> Você vê a obra em consulta. Se você é o líder, combine com o Ighor antes de trocar de novo.");
            aviso("✓ <b>Agora você é o líder</b> de " + esc(o.nome) + ".", "ok"); avisoSome(6000);
            gravar("ultima", o.id); abrirObra(o);
          });
        });
      }).catch(function(e){ throw e; }).then(null, function(e){
        registro("troca de líder falhou: " + (e.message || e));
        aviso("<b>Não consegui trocar o líder.</b> " + esc(e.message || e) + "<br>Nada mudou. Tente de novo com sinal ou avise o Gustavo.", "erro", [{ txt: "OK", fn: function(){ aviso(null); } }]);
        return carregarObras().then(mostrarObras, function(){ var m = ler("obras:" + MODO, null); if(m) mostrarObras(m); });
      });
    });
  });
}
/* ADM: cada líder com as obras em andamento e o histórico; adicionar e remover obra do histórico */
function painelLideres(m, P){
  P.appendChild(tituloGrupo("Líderes"));
  P.appendChild(vazio("O que cada líder vê na aba Minhas obras. Você pode acrescentar ou tirar obras do histórico dele."));
  Nucleo.historicoDosLideres(m, ctxObras()).forEach(function(h){
    var bloco = document.createElement("div"); bloco.className = "lider-bloco";
    var tit = document.createElement("h3"); tit.textContent = h.lider.nome; bloco.appendChild(tit);
    var linhas = [];
    h.minhas.forEach(function(it){ linhas.push({ txt: it.obra.nome, det: "em andamento" + (it.obra.periodo ? " · " + it.obra.periodo : ""), remover: null }); });
    h.realizadas.forEach(function(r){
      var per = r.de ? dataHistorico(r.de) + " a " + (r.ate ? dataHistorico(r.ate) : "hoje") : r.ate ? "até " + dataHistorico(r.ate) : "período não informado";
      linhas.push({ txt: r.obra.nome, det: per + (r.ativa ? " · obra ainda em andamento" : " · encerrada") + (r.manual ? " · incluída pelo ADM" : ""), remover: r.obra });
    });
    if(!linhas.length) bloco.appendChild(vazio("Nenhuma obra."));
    linhas.forEach(function(x){
      var li = document.createElement("div"); li.className = "lider-obra";
      var s = document.createElement("span"); s.innerHTML = "<b></b><small></small>"; s.querySelector("b").textContent = x.txt; s.querySelector("small").textContent = x.det;
      li.appendChild(s);
      if(x.remover){ var b = document.createElement("button"); b.type = "button"; b.className = "mini"; b.textContent = "Remover";
        b.onclick = function(){ removerDoHistorico(m, h.lider, x.remover); }; li.appendChild(b); }
      bloco.appendChild(li);
    });
    var add = document.createElement("button"); add.type = "button"; add.className = "link"; add.textContent = "+ Adicionar obra ao histórico de " + h.lider.nome.split(" ")[0];
    add.onclick = function(){ adicionarAoHistorico(m, h.lider); };
    bloco.appendChild(add);
    P.appendChild(bloco);
  });
}
function gravarAjuste(m, acao, obra, lider, de, ate){
  var t = Nucleo.novoAjuste(acao, obra, lider, de, ate, perfilReal()), f = new File([t.conteudo], t.nome, { type: "application/json" });
  aviso("Gravando…", "");
  return enviarArquivo(f, Nucleo.pastaHistorico(MANIFESTOS[MODO]) + "/" + t.nome, { agora: false, silencioso: true }).then(function(){
    registro("histórico: " + acao + " " + obra.id + " para " + lider.email);
    return carregarExtras().then(function(){ aviso("✓ Histórico de " + esc(lider.nome) + " atualizado.", "ok"); avisoSome(4000); mostrarObras(m); abrirPerfil(); });
  }, function(e){ aviso("<b>Não consegui gravar.</b> " + esc(e.message || e) + "<br>Nada mudou. Tente de novo com sinal.", "erro", [{ txt: "OK", fn: function(){ aviso(null); } }]); });
}
function removerDoHistorico(m, lider, obra){
  confirmar("Tirar do histórico?", obra.nome + "\n\nsai do histórico de " + lider.nome + ". A obra e os envios dela não mudam.", "Tirar", "Cancelar")
    .then(function(ok){ if(ok) gravarAjuste(m, "remover", obra, lider, null, null); });
}
function adicionarAoHistorico(m, lider){
  var f = $("form-hist"), sel = $("fh-obra");
  $("fh-tit").textContent = "Adicionar ao histórico de " + lider.nome;
  sel.replaceChildren();
  [].concat((m.obras || []).map(function(o){ return { o: o, rot: o.nome + " (em andamento)" }; }),
            (m.encerradas || []).map(function(o){ return { o: o, rot: o.nome + " (encerrada" + (o.encerrada_em ? " " + dataHistorico(o.encerrada_em) : "") + ")" }; }))
    .forEach(function(x, i){ var op = document.createElement("option"); op.value = String(i); op.textContent = x.rot; op._o = x.o; sel.appendChild(op); });
  $("fh-de").value = ""; $("fh-ate").value = "";
  f.hidden = false;
  $("fh-cancelar").onclick = function(){ f.hidden = true; };
  $("fh-salvar").onclick = function(){
    var o = sel.options[sel.selectedIndex] && sel.options[sel.selectedIndex]._o; if(!o) return;
    var de = $("fh-de").value || null, ate = $("fh-ate").value || null;
    if(de && ate && ate < de){ aviso("A data final é antes da inicial.", "erro"); avisoSome(3500); return; }
    f.hidden = true;
    gravarAjuste(m, "adicionar", o, lider, de, ate);
  };
}
/* ------------------------------------------------ arquivamento (4 semanas depois do término) */
function depsDaObra(o){
  return buscarPlanta(o).then(function(html){
    return { listar: listarPasta, baixarInicio: baixarInicio, baixarFaixa: baixarFaixa, identidade: html ? Nucleo.identidadeDaPlanta(html) : null };
  });
}
function arquivarAgora(m, o, automatico){
  function fazer(){
    if(!navigator.onLine){ aviso("Sem sinal: para tirar a planta do aparelho, o app confere antes no SharePoint que tudo chegou. Tente com sinal.", "erro"); avisoSome(6000); return Promise.resolve(false); }
    if(!automatico) aviso("Conferindo no SharePoint se tudo desta obra chegou…", "");
    return depsDaObra(o).then(function(deps){ return Nucleo.provarArquivamento(o, deps); }).then(function(r){
      registro("arquivar " + o.id + ": " + r.motivo);
      if(!r.ok){
        if(!automatico) aviso("<b>Não tirei a planta do aparelho.</b> " + esc(r.motivo) + (r.semProva && r.semProva.length ? ":<br><small>" + r.semProva.slice(0, 5).map(esc).join("<br>") + "</small>" : "")
          + "<br>Abra a obra e toque em ENVIAR. Se continuar, avise o Gustavo.", "erro", [{ txt: "OK", fn: function(){ aviso(null); } }]);
        return false;
      }
      return Nucleo.arquivarObra(o).then(function(){
        aviso("✓ <b>" + esc(o.nome) + "</b> saiu deste aparelho. As marcações e fotos estão no SharePoint: no histórico, abra a obra e toque em Sincronizar para ver.", "ok");
        avisoSome(8000); mostrarObras(m); return true;
      });
    }).catch(function(e){ if(!automatico){ aviso("Não deu para conferir agora: " + esc(e.message || e), "erro"); avisoSome(6000); } return false; });
  }
  if(automatico) return fazer();
  return confirmar("Tirar a planta do aparelho?", o.nome + "\n\nO app confere no SharePoint que todas as marcações e fotos chegaram e, só então, apaga a planta deste aparelho. Para ver depois: histórico → Sincronizar.",
    "Tirar do aparelho", "Cancelar").then(function(ok){ if(ok) return fazer(); });
}
function avisosDeArquivo(m, L){
  if(!Nucleo.arquivamentoAtivo()) return;
  Nucleo.obrasParaArquivar(m).forEach(function(x){
    if(x.dias === null) return;
    var o = x.obra, quando = x.prazo.split("-").reverse().join("/");
    if(x.dias <= 0){
      /* passou o prazo: tenta sozinho; se não tiver prova, avisa e não apaga */
      arquivarAgora(m, o, true).then(function(ok){
        if(ok) return;
        var c = cartaoObra(o.nome, "Já passou o prazo para sair deste aparelho, mas falta coisa sem cópia no SharePoint. Abra e toque em ENVIAR.", "Abrir",
          function(){ abrirObra(o, { pendencias: true }); }, "pend");
        L.insertBefore(c, L.firstChild);
      });
      return;
    }
    if(x.dias > 7) return;
    var caixa = document.createElement("div"); caixa.className = "arquivo-aviso";
    var tx = document.createElement("p"); tx.innerHTML = "<b></b><br><small></small>";
    tx.querySelector("b").textContent = o.nome;
    tx.querySelector("small").textContent = "A planta sai deste aparelho em " + quando + " (4 semanas depois do término). As fotos ficam no SharePoint.";
    var b1 = document.createElement("button"); b1.type = "button"; b1.textContent = "Adiar 2 semanas";
    b1.onclick = function(){ var novo = Nucleo.adiarArquivo(o); aviso("Fica no aparelho até " + novo.split("-").reverse().join("/") + ".", "ok"); avisoSome(4000); mostrarObras(m); };
    var b2 = document.createElement("button"); b2.type = "button"; b2.textContent = "Arquivar agora"; b2.className = "g";
    b2.onclick = function(){ arquivarAgora(m, o); };
    var bts = document.createElement("div"); bts.className = "bts"; bts.append(b1, b2);
    caixa.append(tx, bts); L.insertBefore(caixa, L.firstChild);
  });
}
function abrirPerfil(){ $("painel-perfil").hidden = false; $("fundo-perfil").hidden = false; $("b-perfil").setAttribute("aria-expanded", "true"); }
function fecharPerfil(){ var x = $("painel-perfil"); if(!x) return; x.hidden = true; $("fundo-perfil").hidden = true; $("b-perfil").setAttribute("aria-expanded", "false"); }

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
  /* 0.2 (C29): textos da AJUDA trocados pela casca; com espera de 300 ms para não pesar em planta grande */
  + 'var aj=null;function ajus(){clearTimeout(aj);aj=setTimeout(function(){try{P.ajustarTextos(document);}catch(e){}},300);}'
  + 'lim(document.body);try{P.ajustarTextos(document);}catch(e){}'
  + 'new MutationObserver(function(ms){for(var k=0;k<ms.length;k++)lim(ms[k].target);ajus();}).observe(document.body,{childList:true,subtree:true,characterData:true});});'
  + '})();<\/script><style>#env-passos .passo,#env-passos label.campo,#env-corpo ol{display:none!important}</style>';
var OBRA = null;
/* 0.3: faixa com o estado do envio, calculado pelo Nucleo (verde só com o SharePoint confirmando tudo).
   Sem sinal tem prioridade. Na consulta, a faixa diz de quem é a obra e de quando é o envio. */
var CONSULTA = null, PLANTA_INFO = null;
function janelaPlanta(){ try{ return $("planta").contentWindow; }catch(e){ return null; } }
function faixaEnvio(){
  var el = $("faixa-envio");
  if(!OBRA){ exibirFaixa(el, "", ""); return; }
  if(CONSULTA){ exibirFaixa(el, "consulta", CONSULTA.faixa); return; }
  var o = OBRA, info = PLANTA_INFO;
  function atualizar(){
    if(OBRA === o && PLANTA_INFO === info && !CONSULTA) pintarFaixa(el, Nucleo.estadoEnvio(o, info));
  }
  Nucleo.observar(o, info).then(atualizar, atualizar);
}
function exibirFaixa(el, classe, texto){
  el.className = classe;
  if(el.textContent !== texto) el.textContent = texto;
}
function pintarFaixa(el, s){
  var conf = s.confirmadoNoSharePoint ? dataHora(s.confirmadoNoSharePoint) : "";
  var sinal = navigator.onLine ? "" : " · Sem sinal";
  if(s.cor === "azul") exibirFaixa(el, "azul", "Enviando… mantenha o app aberto" + sinal);
  else if(s.cor === "vermelho") exibirFaixa(el, "sem", navigator.onLine
    ? "Envio parou · toque em ENVIAR de novo"
    : "Envio parou · Sem sinal · Toque em ENVIAR quando o sinal voltar");
  else if(s.cor === "verde" && conf) exibirFaixa(el, "ok", "Tudo enviado · " + conf + sinal);
  else if(s.cor === "amarelo"){
    var fotos = s.faltam.fotos;
    var detalhe = [s.ultimaAlteracao ? "alterado " + dataHora(s.ultimaAlteracao) : "",
      fotos ? fotos + (fotos === 1 ? " foto pendente" : " fotos pendentes") : "",
      conf ? "último envio " + conf : ""].filter(Boolean);
    exibirFaixa(el, "nada", ["Não enviado", ...detalhe].join(" · ") + sinal);
  }else if(s.cor === "vazio") exibirFaixa(el, "", "Nada marcado ainda" + sinal);
  else if(s.cor === "migracao") exibirFaixa(el, "", "Último envio pela versão anterior: " + dataHora(s.envioAnterior) + " · toque ENVIAR para conferir" + sinal);
  else{
    var u = ler("ultimoEnvio:" + OBRA.id, null);
    exibirFaixa(el, "", (u ? "Último envio recebido · " + dataHora(u) : "Envio ainda não confirmado") + sinal);
  }
}
window.addEventListener("online", faixaEnvio); window.addEventListener("offline", faixaEnvio);
var AVISOU_TROCA = null, CONFERIU_TROCA = 0;
function conferirTroca(){
  var o = OBRA;
  if(!o || CONSULTA || !navigator.onLine || ENVIANDO || Date.now() - CONFERIU_TROCA < 110000) return;
  CONFERIU_TROCA = Date.now();
  carregarExtras().then(function(){
    if(OBRA !== o || CONSULTA || AVISOU_TROCA === o.id) return;
    var p = perfil(), ed = Nucleo.editorDaObra(o, ctxObras());
    if(!p.papel || !ed || !ed.email || ed.email.toLowerCase() === String(p.email).toLowerCase()) return;
    AVISOU_TROCA = o.id;
    registro("troca percebida com a planta aberta: " + o.id + " agora é de " + ed.email);
    aviso("<b>" + esc(ed.nome || ed.email) + " assumiu esta obra" + (ed.desde ? " às " + hora(ed.desde) : "") + ".</b><br>Envie o que ficou no seu aparelho (ENVIAR) e volte às obras.", "erro",
      [{ txt: "OK", fn: function(){ aviso(null); } }]);
  });
}
setInterval(conferirTroca, 120000);
document.addEventListener("visibilitychange", function(){ if(!document.hidden) conferirTroca(); });
setInterval(faixaEnvio, 30000);
document.addEventListener("visibilitychange", function(){ if(!document.hidden) faixaEnvio(); });

/* 0.3: consulta. Lista da pasta Envios pelo Graph (com "Criado por") e só o começo de cada arquivo (Range). */
/* lista a pasta INTEIRA (o Graph devolve em páginas de até 200: segue o @odata.nextLink) */
function listarPasta(pasta){
  var todos = [];
  function pagina(j){
    (j.value || []).forEach(function(x){ if(x.file) todos.push({ name: x.name, size: x.size, eTag: x.eTag || null, createdDateTime: x.createdDateTime,
                                                                 createdBy: x.createdBy, url: x["@microsoft.graph.downloadUrl"] }); });
    var prox = j["@odata.nextLink"];
    if(!prox || todos.length > 5000) return todos;
    return token().then(function(tk){ return fetch(prox, { headers: { Authorization: "Bearer " + tk } }); })
      .then(function(r){ return r.ok ? r.json() : erroGraph(r); }).then(pagina);
  }
  return drive().then(function(d){ return graph("/drives/" + d + "/root:/" + rel(pasta) + ":/children?$top=200"); }).then(pagina);
}
function baixarInicio(url, n){
  return fetch(url, { headers: { Range: "bytes=0-" + (n - 1) } }).then(function(r){ if(!r.ok) throw new Error("download " + r.status); return r.text(); });
}
function baixarFaixa(url, ini, fim){
  return fetch(url, { headers: { Range: "bytes=" + ini + "-" + fim } }).then(function(r){ if(!r.ok) throw new Error("download " + r.status); return r.text(); });
}
function dataHora(iso){ var d = new Date(iso); return (d.toDateString() === new Date().toDateString() ? "hoje " : d.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" }) + " ") + hora(iso); }
function abrirConsulta(o, it){
  if(ENVIANDO){ aviso("Espere o envio terminar para trocar de obra.", "erro"); avisoSome(4000); return; }
  var resp = (it && it.editor && it.editor.nome) || (Nucleo.editorDaObra(o, ctxObras()) || {}).nome || o.lider || "Não informado";
  function voltar(){ $("b-trocar").onclick(); }
  function indisponivel(){
    mensagem("Consulta indisponível para esta planta", "Se precisar conferir a obra, fale com o responsável.",
      voltar, "Voltar às obras");
  }
  mensagem("Consultando " + o.nome + "…", "Buscando o último envio confirmado na pasta da obra.");
  var ant = CONSULTA;
  /* a pasta Envios pode ser dividida entre obras (ex.: _TESTE PLANTA/Envios): só vale envio da MESMA planta */
  buscarPlanta(o).then(function(html){
    if(!html) return [html, { achou: false }];
    return Nucleo.ultimoEnvioConfirmado(o, { listar: listarPasta, baixarInicio: baixarInicio, baixarFaixa: baixarFaixa, identidade: Nucleo.identidadeDaPlanta(html) })
      .then(function(env){ return [html, env]; });
  }).then(function(x){
    var html = x[0], env = x[1];
    if(!html){ mensagem("Não deu para abrir a planta", "Precisa de sinal na primeira vez.", function(){ abrirConsulta(o, it); }); return; }
    if(Nucleo.modoDaObra(o, perfil(), ctxObras(), html).modo === "indisponivel"){ indisponivel(); return; }
    if(!env.achou){
      if(o.encerrada_em && env.situacao !== "so_incompleto"){
        var c0 = Nucleo.htmlDeConsulta(html, { reg: {}, oc: [], lider: "" });
        if(c0) return (ant ? descarregarPlanta().then(function(){ return Nucleo.limparConsulta(ant); }) : Promise.resolve()).then(function(){
          OBRA = o;
          CONSULTA = { obra: o.id, chave: c0.chave, banco: c0.banco, envio: null, identidade: Nucleo.identidadeDaPlanta(html), encerrada: true,
            faixa: "Só consulta · " + o.nome.replace(/^TESTE · /, "") + " · nenhum envio desta planta foi achado" };
          tela(null); $("faixa-nome").textContent = (MODO === "teste" ? "TESTE · " : "") + o.nome.replace(/^TESTE · /, "");
          PLANTA_INFO = null; $("planta").onload = null; faixaEnvio(); $("planta").srcdoc = c0.html; atualizarBotaoFotos();
        });
      }
      if(env.situacao === "so_incompleto")
        mensagem("O envio ainda não chegou inteiro", "Chegou só parte do último envio (" + env.incompletoMaisNovo.partes + " partes). As marcações aparecem quando todas as partes chegarem.",
          voltar, "Voltar às obras");
      else
        mensagem("Ainda não há envio para consultar", "As marcações aparecem aqui após o primeiro envio completo.",
          voltar, "Voltar às obras");
      return;
    }
    var c = Nucleo.htmlDeConsulta(html, env);
    if(!c){ indisponivel(); return; }
    return (ant ? descarregarPlanta().then(function(){ return Nucleo.limparConsulta(ant); }) : Promise.resolve()).then(function(){
      OBRA = o;
      CONSULTA = { obra: o.id, chave: c.chave, banco: c.banco, envio: env, identidade: Nucleo.identidadeDaPlanta(html), encerrada: !!o.encerrada_em,
        faixa: "Só consulta · Responsável: " + resp + " · Último envio confirmado: " + dataHora(env.confirmadoEm) + (env.autor ? " · Enviado por " + env.autor : "") };
      tela(null); document.title = "Consulta · " + o.nome;
      $("faixa-nome").textContent = (MODO === "teste" ? "TESTE · " : "") + o.nome.replace(/^TESTE · /, "");
      PLANTA_INFO = null; $("planta").onload = null;
      faixaEnvio();
      $("planta").srcdoc = c.html;
      atualizarBotaoFotos();
      registro("consulta " + o.id + " envio " + env.arquivo + " de " + (env.autor || "?"));
    });
  }).catch(function(e){
    registro("consulta " + o.id + ": " + (e.message || e));
    mensagem("Não deu para consultar", "Confira a conexão e tente de novo.", function(){ abrirConsulta(o, it); });
  });
}
/* a planta da consulta segura o banco de fotos aberto: descarrega o iframe antes de apagar o que ela deixou */
/* "Ver como" (teste): a planta do líder para tocar, marcar e tirar foto, isolada e sem envio */
function abrirSimulacao(o, p){
  if(ENVIANDO){ aviso("Espere o envio terminar para trocar de obra.", "erro"); avisoSome(4000); return; }
  mensagem("Abrindo a demonstração de " + o.nome + "…", "");
  var ant = CONSULTA;
  buscarPlanta(o).then(function(html){
    var c = html && Nucleo.htmlDeSimulacao(html);
    if(!c){ mensagem("Demonstração indisponível para esta planta", "Planta por módulo ou de remanejamento: abra pela conta do líder no teste.", function(){ $("b-trocar").onclick(); }, "Voltar às obras"); return; }
    return (ant ? descarregarPlanta().then(function(){ return Nucleo.limparConsulta(ant); }) : Promise.resolve()).then(function(){
      OBRA = o;
      CONSULTA = { obra: o.id, chave: c.chave, banco: c.banco, simulacao: true,
        faixa: "Demonstração · como " + p.nome + " vê · nada é enviado nem guardado" };
      tela(null); document.title = "Demonstração · " + o.nome;
      $("faixa-nome").textContent = "TESTE · " + o.nome.replace(/^TESTE · /, "");
      PLANTA_INFO = null; $("planta").onload = null;
      faixaEnvio(); $("b-fotos").hidden = true;
      $("planta").srcdoc = c.html;
      registro("demonstração " + o.id + " como " + p.email);
    });
  });
}
Nucleo._avisoSimulacao = function(){ aviso("Demonstração: o ENVIAR não envia nada. Para enviar de verdade, entre com a conta do líder.", ""); avisoSome(5000); };
function descarregarPlanta(){
  return new Promise(function(ok){ var fr = $("planta"); fr.onload = function(){ fr.onload = null; ok(); };
    fr.removeAttribute("srcdoc"); fr.src = "about:blank"; setTimeout(ok, 1500); });
}
function sairDaConsulta(){ var c = CONSULTA; CONSULTA = null; fecharGaleria(); $("b-fotos").hidden = true;
  if(c && c.encerrada) Nucleo.apagarFotosSincronizadas(c.obra).then(function(n){ if(n) registro("fotos de " + c.obra + " (histórico) apagadas do aparelho: " + n); }); if(!c) return Promise.resolve(); return descarregarPlanta().then(function(){ return Nucleo.limparConsulta(c); }); }
/* 0.3 (Gustavo, 18/11): Sincronizar fotos. As fotos vêm dos envios da obra no SharePoint e ficam guardadas
   neste aparelho (consulta e histórico). Nada aqui marca nem envia. */
var GALERIA_URLS = [];
function nomePeca(id){ return String(id || "sem peça").replace(/^peca-/, ""); }
function atualizarBotaoFotos(){
  var b = $("b-fotos"); if(!CONSULTA || !OBRA || CONSULTA.simulacao){ b.hidden = true; return; }
  var o = OBRA; b.hidden = false;
  Nucleo.fotosSincronizadas(o.id).then(function(s){
    if(OBRA !== o || !CONSULTA) return;
    b.textContent = s.fotos.length ? "Fotos (" + s.fotos.length + ")" : "Sincronizar fotos";
  });
}
function sincronizarFotosDaObra(){
  var o = OBRA; if(!o || !CONSULTA) return;
  if(!navigator.onLine){ aviso("Sem sinal. As fotos já sincronizadas continuam aqui; para buscar as novas, precisa de sinal.", "erro"); avisoSome(6000); return; }
  aviso("Procurando fotos nos envios de " + esc(o.nome) + "…", "");
  var sessaoC = CONSULTA;
  var deps = { listar: listarPasta, baixarInicio: baixarInicio, identidade: CONSULTA.identidade,
               cancelado: function(){ return CONSULTA !== sessaoC; },
               baixarTudo: function(url){ return fetch(url).then(function(r){ if(!r.ok) throw new Error("download " + r.status); return r.text(); }); } };
  Nucleo.planoDeSincronia(o, deps).then(function(plano){
    if(!plano.novos.length){ aviso("✓ As fotos já estão em dia neste aparelho.", "ok"); avisoSome(4000); abrirGaleria(null); return; }
    var mbs = mb(plano.bytes);
    var pergunta = plano.bytes > 15 * 1024 * 1024
      ? confirmar("Baixar as fotos?", plano.novos.length + " envio(s), " + mbs + " MB. Use Wi-Fi se puder.", "Baixar", "Agora não")
      : Promise.resolve(true);
    return pergunta.then(function(sim){
      if(!sim){ aviso(null); return; }
      return Nucleo.sincronizarFotos(o, plano, deps, function(i, n, item){
        aviso("Sincronizando fotos<br>envio " + (i + 1) + " de " + n + " · " + mb(item.size) + " MB", "", null, i / n);
      }).then(function(r){
        if(r.cancelado || CONSULTA !== sessaoC){
          if(sessaoC.encerrada) Nucleo.apagarFotosSincronizadas(sessaoC.obra);
          registro("fotos " + o.id + ": sincronização cancelada (saiu da consulta)"); aviso(null); return;
        }
        registro("fotos " + o.id + ": " + r.fotos + " foto(s) de " + r.envios + " envio(s)" + (r.falhas.length ? ", " + r.falhas.length + " falha(s)" : ""));
        var ruins = r.falhas.filter(function(x){ return x.tipo === "arquivo_ruim"; }).length, rede = r.falhas.length - ruins;
        if(r.falhas.length) aviso("Sincronizei " + r.fotos + " foto(s)." + (rede ? " " + rede + " envio(s) não baixaram: toque em Sincronizar de novo com sinal melhor." : "")
          + (ruins ? " " + ruins + " arquivo(s) da pasta estão estragados e foram pulados: avise o Gustavo." : ""), "erro", [{ txt: "OK", fn: function(){ aviso(null); } }]);
        else { aviso("✓ " + r.fotos + " foto(s) nova(s) sincronizada(s).", "ok"); avisoSome(4000); }
        atualizarBotaoFotos(); abrirGaleria(null);
      });
    });
  }).catch(function(e){
    registro("fotos " + o.id + ": " + (e.message || e));
    aviso("<b>Não consegui sincronizar as fotos.</b> " + esc(e.message || e) + "<br>Confira o sinal e tente de novo.", "erro", [{ txt: "OK", fn: function(){ aviso(null); } }]);
  });
}
function fecharGaleria(){ $("galeria").hidden = true; $("foto-grande").hidden = true; GALERIA_URLS.forEach(function(u){ URL.revokeObjectURL(u); }); GALERIA_URLS = []; }
function abrirGaleria(pecaId){
  var o = OBRA; if(!o) return;
  Nucleo.fotosSincronizadas(o.id).then(function(s){
    fecharGaleria();
    $("gal-tit").textContent = pecaId ? "Fotos · " + nomePeca(pecaId) : "Fotos · " + o.nome.replace(/^TESTE · /, "");
    var info = s.fotos.length
      ? s.fotos.length + " foto(s) de " + s.envios.length + " envio(s) · sincronizado " + dataHora(s.sincronizadoEm)
        + (s.ultimoEnvio ? "\nÚltimo envio com fotos: " + dataHora(s.ultimoEnvio) + (s.autor ? " · " + s.autor : "") : "")
      : "Nenhuma foto neste aparelho ainda. Toque em Sincronizar fotos.";
    $("gal-info").textContent = info;
    $("gal-todas").hidden = !pecaId;
    var G = $("gal-grade"); G.replaceChildren();
    var lista = pecaId ? s.fotos.filter(function(f){ return f.peca === pecaId; }) : s.fotos;
    if(pecaId && !lista.length){ var v = document.createElement("p"); v.className = "vazio"; v.textContent = "Esta peça não tem foto nos envios sincronizados."; G.appendChild(v); }
    var atual = null;
    lista.forEach(function(f){
      if(f.peca !== atual){ atual = f.peca; var h = document.createElement("h3"); h.textContent = nomePeca(f.peca); G.appendChild(h); }
      var u = URL.createObjectURL(f.blob); GALERIA_URLS.push(u);
      var fig = document.createElement("figure"), im = document.createElement("img"), cap = document.createElement("figcaption");
      im.src = u; im.alt = nomePeca(f.peca) + " " + (f.alvo || ""); im.loading = "lazy";
      cap.textContent = (f.alvo === "ini" ? "antes" : f.alvo === "fim" ? "depois" : (f.alvo || "")) + (f.quando ? " · " + dataHora(f.quando) : "");
      fig.append(im, cap); fig.onclick = function(){ $("foto-grande-img").src = u; $("foto-grande-txt").textContent = nomePeca(f.peca) + " · " + cap.textContent; $("foto-grande").hidden = false; };
      G.appendChild(fig);
    });
    $("galeria").hidden = false;
  });
}
Nucleo._avisoConsulta = function(pecaId){
  if(pecaId && OBRA){
    var o = OBRA;
    Nucleo.fotosSincronizadas(o.id).then(function(s){
      if(s.fotos.some(function(f){ return f.peca === pecaId; })){ abrirGaleria(pecaId); return; }
      aviso(s.fotos.length ? "Só consulta · esta peça não tem foto nos envios sincronizados."
                           : "Só consulta · para ver as fotos, toque em <b>Sincronizar fotos</b> lá em cima.", "");
      avisoSome(4500);
    });
    return;
  }
  aviso("Só consulta. Se encontrar algo para corrigir, fale com o responsável pela obra.", "");
  avisoSome(4500);
};
function abrirObra(o, opc){
  opc = opc || {};
  $("b-fotos").hidden = true;
  if(CONSULTA){ sairDaConsulta().then(function(){ abrirObra(o, opc); }); return; }
  AVISOU_TROCA = null;
  if(ENVIANDO && OBRA && OBRA.id !== o.id){ aviso("Espere o envio terminar para trocar de obra.", "erro"); avisoSome(4000); return; }
  OBRA = o;
  mensagem("Abrindo " + o.nome + "…", navigator.onLine ? "Baixando a planta da pasta da obra." : "Sem sinal: abrindo a última planta guardada.");
  buscarPlanta(o).then(function(html){
    if(!html){ mensagem("Não deu para abrir a planta", "Este aparelho ainda não tem a planta guardada e não conseguiu baixar. Precisa de sinal na primeira vez.", function(){ abrirObra(o); }); return; }
    var i = html.search(/<head[^>]*>/i), fim = i < 0 ? 0 : i + html.match(/<head[^>]*>/i)[0].length;
    tela(null); document.title = o.nome + " · Planta de Campo";
    $("faixa-nome").textContent = (MODO === "teste" ? "TESTE · " : "") + o.nome.replace(/^TESTE · /, "");
    PLANTA_INFO = Nucleo.infoDaPlanta(html); Nucleo.lembrarPlanta(o.id, PLANTA_INFO);
    Nucleo.nomeDoLiderNaPlanta(o, perfil(), ctxObras(), PLANTA_INFO);
    faixaEnvio();
    $("planta").onload = function(){ setTimeout(faixaEnvio, 300); };
    $("planta").srcdoc = html.slice(0, fim) + INJETAR + html.slice(fim);
    registro("abriu " + o.id + (opc.pendencias ? " (só para enviar o que ficou)" : ""));
    if(opc.pendencias){ AVISOU_TROCA = o.id;
      aviso("<b>" + esc(opc.lider || "Outra pessoa") + " é o líder desta obra agora.</b><br>Envie o que ficou no seu aparelho (ENVIAR) e não marque nada novo.", "erro",
        [{ txt: "OK", fn: function(){ aviso(null); } }]); }
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
  /* 0.2 (C09): sem prometer "nada foi perdido"; o que é verdade é que a planta continua guardada no aparelho */
  aviso("<b>Sem sinal agora.</b> O envio sai sozinho quando o sinal voltar. Deixe a planta aberta: o que você marcou continua guardado neste aparelho.", "erro");
  return new Promise(function(ok){ window.addEventListener("online", function f(){ window.removeEventListener("online", f); ok(); }); });
}
/* 0.2 (D12c): conteúdo do arquivo que já está na Envios, para comparar com o que o aparelho quer mandar */
function baixarBytes(p){
  var d0;
  return drive().then(function(d){ d0 = d; return graph("/drives/" + d + "/root:/" + rel(p)); }).then(function(it){
    var u = it["@microsoft.graph.downloadUrl"];
    var r = u ? fetch(u) : token().then(function(tk){ return fetch("https://graph.microsoft.com/v1.0/drives/" + d0 + "/items/" + it.id + "/content", { headers: { Authorization: "Bearer " + tk } }); });
    return r.then(function(x){ if(!x.ok) throw new Error("download " + x.status); return x.arrayBuffer(); });
  });
}
function igual(a, b){
  if(a.byteLength !== b.byteLength) return false;
  var x = new Uint8Array(a), y = new Uint8Array(b);
  for(var i = 0; i < x.length; i++) if(x[i] !== y[i]) return false;
  return true;
}
/* PASTA/NOME.txt -> PASTA/NOME-2.txt (o começo do nome não muda: os importadores procuram pelo prefixo) */
function outroNome(p, n){
  var i = p.lastIndexOf("."), barra = p.lastIndexOf("/");
  return (i > barra ? p.slice(0, i) + "-" + n + p.slice(i) : p + "-" + n);
}
function enviarArquivo(f, destino, desistir){
  var total = f.size, url = null, tentativa = 0, original = destino, nTentNome = 1;
  function sessao(){
    return drive().then(function(d){
      return graph("/drives/" + d + "/root:/" + rel(destino) + ":/createUploadSession",
        { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "fail" } }) });
    }).then(function(j){ url = j.uploadUrl; });
  }
  /* já existe com o mesmo nome. 0.2 (05/11, laboratório D12c): tamanho igual não prova que é o mesmo arquivo.
     Compara o conteúdo: igual = toque repetido (já chegou); diferente = envia com outro nome (-2, -3...),
     sem sobrescrever o que está lá e sem dizer "chegou" para um arquivo que não subiu. */
  function jaExiste(){
    return baixarBytes(destino).then(function(lá){
      return f.arrayBuffer().then(function(aqui){
        if(igual(lá, aqui)) return { size: total, repetido: true };
        nTentNome++;
        if(nTentNome > 9) throw new Error("já existem arquivos com esse nome e outro conteúdo");
        var novo = outroNome(original, nTentNome);
        registro("envio " + f.name + ": já existe outro arquivo com o mesmo nome; enviando como " + novo.split("/").pop());
        destino = novo; url = null;
        return sessao().then(function(){ return pedaco(0); }, function(e){
          if(e.status === 409 || e.codigo === "nameAlreadyExists") return jaExiste();
          throw e;
        });
      });
    });
  }
  function pedaco(ini){
    var fim = Math.min(ini + PEDACO, total);
    if(!desistir.silencioso) aviso("Enviando <b>" + f.name.replace(/</g, "&lt;") + "</b><br>" + mb(ini) + " de " + mb(total) + " MB", "", null, total ? ini / total : 0);
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
          if(desistir.silencioso && tentativa > 2) throw new Error("sem sinal para gravar a troca");
          var s = Math.min(60, 2 * Math.pow(2, Math.min(tentativa, 5)));
          registro("envio " + f.name + ": " + (e.message || e) + " · nova tentativa em " + s + "s");
          if(!desistir.silencioso) aviso("A rede caiu no meio do envio. Tentando de novo em " + s + " s…", "erro",
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
      /* a sessão acabou: ou o arquivo já chegou (resposta perdida) ou expirou. Nome já existe = compara o conteúdo. */
      return sessao().then(function(){ return pedaco(0); }, function(e){
        if(e.status === 409 || e.codigo === "nameAlreadyExists") return jaExiste();
        throw e;
      });
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
var FILA = Promise.resolve(), NA_FILA = {};
/* 0.2 (C29): na AJUDA das plantas, o que só vale para a planta por arquivo (abrir pelo Chrome/OneDrive,
   atalho da Envios, escolher a pasta no ENVIAR, RECUPERAR do arquivo enviado) vira o texto do online.
   O plano B (Downloads + WhatsApp como Documento) continua: é o que vale quando o envio falha.
   Roda pela casca sobre o documento da planta: o HTML da obra não muda. */
var TEXTOS_ONLINE = [
  { el: "div", re: /^\s*ANTES DO 1º DIA/, html: "<b>No app online:</b> entre com a sua conta Montaggio e escolha a obra. O <b>ENVIAR</b> grava direto na pasta Envios da obra e só mostra ✓ depois que o SharePoint confirmou." },
  { el: "li", re: /^\s*A planta abriu em branco/, esconder: true },
  { el: "li", re: /^\s*O ENVIAR deu erro:/, html: "<b>O ENVIAR deu erro:</b> a planta tenta salvar um arquivo em <b>Downloads</b>. Confira se ele apareceu lá e mande ao Gustavo pelo WhatsApp <b>como Documento</b>. Tire print do erro." },
  { el: "li", re: /^\s*A planta abriu vazia:/, html: "<b>A planta abriu vazia:</b> não marque nada de novo. Avise o Gustavo: ele devolve a planta a partir do seu último envio." },
  { el: "li", re: /^\s*Fim do dia: ENVIAR → OneDrive/, html: "Fim do dia: <b>ENVIAR</b>. O arquivo vai direto para a pasta Envios da obra. Sem foto também envia." },
  { el: "ol", re: /role para baixo.*OneDrive/, esconder: true }
];
function ajustarTextos(doc){
  TEXTOS_ONLINE.forEach(function(r){
    var l = doc.querySelectorAll(r.el + ":not([data-online])");
    for(var i = 0; i < l.length; i++){
      if(!r.re.test(l[i].textContent)) continue;
      l[i].setAttribute("data-online", "1");
      if(r.esconder) l[i].style.display = "none"; else l[i].innerHTML = r.html;
    }
  });
}
window.PlantaOnline = {
  versao: VERSAO,
  ajustarTextos: ajustarTextos,
  enviar: function(arqs){
    var o = OBRA, desistir = { agora: false };
    if(CONSULTA) return Promise.reject(new DOMException("Só consulta", "AbortError"));
    if(!o || !arqs || !arqs.length) return Promise.reject(new Error("obra não definida"));
    var chaveFila = o.id + "|" + arqs.map(function(f){ return f.name + ":" + f.size; }).join(",");
    if(NA_FILA[chaveFila]) return Promise.reject(new DOMException("Este arquivo já está sendo enviado", "AbortError"));   /* toque repetido */
    CONFERIU_TROCA = 0; conferirTroca();   /* outro líder assumiu? avisa (não bloqueia o envio do que já foi feito) */
    ENVIANDO++;
    Nucleo.envioComecou(o, arqs); faixaEnvio();
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
      var pt = Nucleo.parteDoNome(arqs[0].name);
      if(pt && pt.parte < pt.partes)
        aviso("✓ <b>Parte " + pt.parte + " de " + pt.partes + " recebida</b> às " + hora(new Date()) + ".<br>Toque em ENVIAR de novo para mandar a parte " + (pt.parte + 1) + ".", "ok");
      else
        aviso("✓ <b>Chegou no SharePoint</b> às " + hora(new Date()) + "<br>" + (pt ? "Última parte (" + pt.partes + " de " + pt.partes + "), " : arqs.length + " arquivo(s), ") + mb(b) + " MB, na pasta Envios de " + o.nome.replace(/</g, "&lt;") + ".", "ok");
      avisoSome(9000);
      gravar("ultimoEnvio:" + o.id, new Date().toISOString());
      Nucleo.envioTerminou(o, arqs, true, PLANTA_INFO).then(function(){ faixaEnvio(); setTimeout(faixaEnvio, 700); setTimeout(faixaEnvio, 2000); });
    }, function(e){
      Nucleo.envioTerminou(o, arqs, false, PLANTA_INFO).then(faixaEnvio);
      registro("falhou: " + (e.message || e));
      aviso("<b>Não chegou no SharePoint.</b> " + esc(e.message || e) + "<br>A planta vai tentar salvar o arquivo em <b>Downloads</b>. Confira se ele apareceu lá e mande ao Gustavo pelo WhatsApp <b>como Documento</b>. Não apareceu? Não apague nada e ligue para o Gustavo.", "erro",
        [{ txt: "OK", fn: function(){ aviso(null); } }]);
      var x = new Error(e.message || String(e)); x.name = "EnvioSharePoint"; throw x;
    });
    NA_FILA[chaveFila] = p;
    FILA = p.catch(function(){});
    FILA.then(function(){ ENVIANDO--; delete NA_FILA[chaveFila]; });
    return p;
  }
};

/* ------------------------------------------------ partida */
$("b-entrar").onclick = function(){ pca.loginRedirect({ scopes: ESCOPOS }); };
/* trocar de obra: a planta aberta já guardou tudo no aparelho; só não troca no meio de um envio */
var ENVIANDO = 0;
/* 0.2 (C05): a planta some da tela no mesmo toque (não dá para começar um envio enquanto a lista carrega);
   e abrirObra confere de novo antes de trocar a planta */
$("b-trocar").onclick = function(){
  if(CONSULTA){ sairDaConsulta().then(function(){ $("b-trocar").onclick(); }); return; }
  if(ENVIANDO){ aviso("Espere o envio terminar para trocar de obra.", "erro"); avisoSome(4000); return; }
  mensagem("Carregando as obras…", "");
  carregarObras().then(function(m){ return carregarExtras().then(function(){ mostrarObras(m); }); }).catch(function(){ var m = ler("obras:" + MODO, null); if(m) mostrarObras(m); else if(OBRA) abrirObra(OBRA); });
};
$("b-fotos").onclick = function(){ if(this.textContent.indexOf("Fotos (") === 0) abrirGaleria(null); else sincronizarFotosDaObra(); };
$("gal-sinc").onclick = sincronizarFotosDaObra; $("gal-fechar").onclick = fecharGaleria; $("gal-todas").onclick = function(){ abrirGaleria(null); };
$("foto-grande").onclick = function(){ $("foto-grande").hidden = true; };
$("b-perfil").onclick = abrirPerfil; $("b-fechar-perfil").onclick = fecharPerfil; $("fundo-perfil").onclick = fecharPerfil;
$("b-sair").onclick = function(){ pca.logoutRedirect({ account: conta(), postLogoutRedirectUri: BASE }); };
function partida(){
  if(!conta()){ tela("t-entrar"); return; }
  mensagem("Carregando…", "");
  carregarObras().then(function(m){
    var u = ler("ultima", null), o = (m.obras || []).filter(function(x){ return x.id === u; })[0];
    return carregarExtras().then(function(){
      var minha = o && Nucleo.listasDaConta(m, perfil(), ctxObras()).minhas.some(function(it){ return it.obra.id === o.id; });
      if(minha && !/obras/.test(location.hash)) abrirObra(o); else mostrarObras(m); });
  }).catch(function(e){
    /* 404: a lista não existe (ainda não publicada); 403: a conta do líder não tem acesso ao site Montagem */
    var txt = e.status === 404 ? "A lista de obras ainda não foi publicada para este modo (" + MODO + "). Avise o Gustavo."
      : e.status === 403 ? "A sua conta ainda não tem acesso às pastas das obras no SharePoint. Avise o Gustavo: ele pede a liberação."
      : "Precisa de sinal na primeira vez.";
    mensagem("Não deu para ler a lista de obras", txt + "<br><small>Detalhe: " + String(e.message || e).replace(/</g, "&lt;") + "</small>", partida);
  });
}
/* 0.2 (C13): pede ao navegador para não apagar a planta e as fotos quando o aparelho fica sem espaço,
   e avisa se o espaço está acabando. O Chrome pode negar o pedido; isso não protege contra "limpar dados". */
function protegerArmazenamento(){
  var s = navigator.storage; if(!s || !s.persist) return;
  s.persisted().then(function(p){ return p || s.persist(); }).then(function(p){
    registro("armazenamento protegido: " + (p ? "sim" : "não"));
    return s.estimate ? s.estimate() : null;
  }).then(function(e){
    if(!e || !e.quota) return;
    var livre = e.quota - (e.usage || 0);
    registro("espaço: usado " + mb(e.usage || 0) + " MB de " + mb(e.quota) + " MB");
    if(livre < 300 * 1048576 || (e.usage || 0) / e.quota > 0.85)
      aviso("<b>Pouco espaço livre neste aparelho</b> (" + mb(livre) + " MB). Envie o que está pendente e avise o Gustavo antes de tirar mais fotos.", "erro", [{ txt: "OK", fn: function(){ aviso(null); } }]);
  }).catch(function(e){ registro("armazenamento: " + (e.message || e)); });
}
protegerArmazenamento();
pca.initialize().then(function(){ return pca.handleRedirectPromise(); }).then(partida).catch(function(e){
  registro("login: " + (e.errorCode || "") + " " + (e.message || e));
  tela("t-entrar"); aviso("O login não terminou: " + esc(e.errorMessage || e.message || e), "erro", [{ txt: "OK", fn: function(){ aviso(null); } }]);
});
if("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(function(e){ registro("sw: " + e.message); });
