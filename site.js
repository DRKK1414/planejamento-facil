// Planejamento Fácil: site onde a professora escreve e gera o Word no modelo que a extensão lê.
// Tudo fica no navegador dela (localStorage). Nada é enviado para servidor nenhum.
(() => {
  const CHAVE = 'planejamento-facil-dados';
  const $ = (s, r = document) => r.querySelector(s);
  const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const novoId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

  // ---------- Dados ----------
  let dados = { planejamentos: [], diario: [], especial: [], ultimo: {} };
  try { const s = JSON.parse(localStorage.getItem(CHAVE)); if (s && typeof s === 'object') dados = { ...dados, ...s }; } catch (e) { }
  let avisoSemGuardar = false;
  function guardar() {
    try { localStorage.setItem(CHAVE, JSON.stringify(dados)); return true; }
    catch (e) { avisoSemGuardar = true; return false; }
  }

  // ---------- Datas ----------
  const hojeISO = () => { const d = new Date(); return iso(d); };
  const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const deISO = s => { const [a, m, d] = s.split('-').map(Number); return new Date(a, m - 1, d); };
  const br = s => s ? s.split('-').reverse().join('/') : '';
  const brCurto = s => br(s).slice(0, 5);
  const DIAS = ['domingo', 'segunda', 'terça', 'quarta', 'quinta', 'sexta', 'sábado'];
  const segunda = s => { const d = deISO(s); d.setDate(d.getDate() - ((d.getDay() + 6) % 7)); return iso(d); };
  const somaDias = (s, n) => { const d = deISO(s); d.setDate(d.getDate() + n); return iso(d); };

  // ---------- Word (.docx) sem bibliotecas ----------
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  function zip(arquivos) {
    const enc = new TextEncoder(), partes = [], central = [];
    let pos = 0;
    for (const [nome, texto] of arquivos) {
      const n = enc.encode(nome), d = enc.encode(texto), c = crc32(d);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
      h.setUint16(10, 0, true); h.setUint16(12, 0x21, true); h.setUint32(14, c, true); h.setUint32(18, d.length, true);
      h.setUint32(22, d.length, true); h.setUint16(26, n.length, true); h.setUint16(28, 0, true);
      partes.push(new Uint8Array(h.buffer), n, d);
      const cd = new DataView(new ArrayBuffer(46));
      cd.setUint32(0, 0x02014b50, true); cd.setUint16(4, 20, true); cd.setUint16(6, 20, true); cd.setUint16(8, 0x0800, true);
      cd.setUint16(10, 0, true); cd.setUint16(12, 0, true); cd.setUint16(14, 0x21, true); cd.setUint32(16, c, true);
      cd.setUint32(20, d.length, true); cd.setUint32(24, d.length, true); cd.setUint16(28, n.length, true);
      cd.setUint32(42, pos, true);
      central.push(new Uint8Array(cd.buffer), n);
      pos += 30 + n.length + d.length;
    }
    const tamCentral = central.reduce((s, p) => s + p.length, 0);
    const fim = new DataView(new ArrayBuffer(22));
    fim.setUint32(0, 0x06054b50, true); fim.setUint16(8, arquivos.length, true); fim.setUint16(10, arquivos.length, true);
    fim.setUint32(12, tamCentral, true); fim.setUint32(16, pos, true);
    return new Blob([...partes, ...central, new Uint8Array(fim.buffer)], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }

  const x = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  // Um parágrafo. o: {b, i, tam (pt), cor, centro, depois (espaço em pt)}
  function par(texto, o = {}) {
    const rpr = (o.b ? '<w:b/>' : '') + (o.i ? '<w:i/>' : '') + (o.cor ? `<w:color w:val="${o.cor}"/>` : '') + (o.tam ? `<w:sz w:val="${o.tam * 2}"/>` : '');
    const ppr = (o.centro ? '<w:jc w:val="center"/>' : '') + `<w:spacing w:before="${(o.antes || 0) * 20}" w:after="${(o.depois ?? 4) * 20}"/>`;
    const run = texto ? `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${x(texto)}</w:t></w:r>` : '';
    return `<w:p><w:pPr>${ppr}</w:pPr>${run}</w:p>`;
  }
  const linhas = (texto, o) => { const ls = String(texto || '').split('\n'); return ls.map(l => par(l, o)).join(''); };
  const dica = t => par(t, { i: true, cor: '808080', tam: 9, depois: 6 });
  const titulo = t => par(t, { b: true, tam: 14, cor: '1F4E79', antes: 12, depois: 4 });
  function celula(texto, largura, o = {}) {
    const fundo = o.fundo ? `<w:shd w:val="clear" w:color="auto" w:fill="${o.fundo}"/>` : '';
    return `<w:tc><w:tcPr><w:tcW w:w="${largura}" w:type="dxa"/>${fundo}</w:tcPr>${linhas(texto, { b: o.b, depois: 0 })}</w:tc>`;
  }
  function tabela(larguras, linhasTab, cabecalho) {
    const bordas = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(l => `<w:${l} w:val="single" w:sz="4" w:space="0" w:color="8C9AA8"/>`).join('');
    const grade = larguras.map(w => `<w:gridCol w:w="${w}"/>`).join('');
    const trs = linhasTab.map((ln, i) => `<w:tr>${ln.map((t, j) => celula(t, larguras[j], (cabecalho && i === 0) ? { b: true, fundo: 'DCE6F0' } : (!cabecalho && j === 0 ? { b: true } : {}))).join('')}</w:tr>`).join('');
    return `<w:tbl><w:tblPr><w:tblW w:w="${larguras.reduce((a, b) => a + b, 0)}" w:type="dxa"/><w:tblBorders>${bordas}</w:tblBorders><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${grade}</w:tblGrid>${trs}</w:tbl>`;
  }
  function docx(corpo, paisagem) {
    const pg = paisagem ? '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="850" w:right="850" w:bottom="850" w:left="850" w:header="0" w:footer="0" w:gutter="0"/>'
      : '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>';
    const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${corpo}<w:sectPr>${pg}</w:sectPr></w:body></w:document>`;
    const estilos = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/><w:sz w:val="22"/><w:lang w:val="pt-BR"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;
    return zip([
      ['[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'],
      ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'],
      ['word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'],
      ['word/document.xml', doc],
      ['word/styles.xml', estilos],
    ]);
  }

  // Os títulos e a ordem são os que a extensão procura (não mudar)
  const SECOES = [
    ['objetivos', 'Objetivos', ''],
    ['conteudo', 'Conteúdo', ''],
    ['estrategia', 'Estratégia', ''],
    ['instrumentos', 'Instrumentos', 'Um por linha, com o mesmo nome da lista do sistema (ex.: Leitura, Pesquisa).'],
    ['criterios', 'Critérios de avaliação', 'Um por linha, com o mesmo nome da lista do sistema.'],
    ['referencias', 'Referências', ''],
    ['objetos', 'Objetos de Conhecimento', 'Um por linha: o código do sistema ou o nome do objeto.'],
    ['habilidades', 'Habilidades', 'Um código BNCC por linha (ex.: EF06GE11).'],
  ];

  function wordPlanejamento(p) {
    let c = par('PLANEJAMENTO', { b: true, tam: 20, centro: true, cor: '1F4E79', depois: 4 });
    c += dica('Gerado pelo site Planejamento Fácil. Não mude os títulos em azul nem a primeira coluna da tabela: a extensão usa esses títulos para saber onde colar cada parte.');
    c += titulo('Cabeçalho');
    c += tabela([2800, 6800], [['Turma', p.turma], ['Disciplina', p.disciplina], ['Professor(a)', p.professor || ''], ['Trimestre', p.trimestre], ['Tema/Projeto', p.tema],
      ['Data início', br(p.inicio) || 'dd/mm/aaaa'], ['Data término', br(p.termino) || 'dd/mm/aaaa']], false);
    for (const [k, nome, d] of SECOES) {
      c += titulo(nome);
      if (d) c += dica(d);
      c += linhas(p[k] || '');
    }
    return docx(c, false);
  }

  function wordDiario(tipo, itens) {
    const turmas = [...new Set(itens.map(i => i.turma || ''))];
    let c = par(tipo === 'especial' ? 'DIÁRIO DA SEMANA · EDUCAÇÃO ESPECIAL' : 'DIÁRIO DA SEMANA', { b: true, tam: 20, centro: true, cor: '1F4E79', depois: 4 });
    c += dica('Gerado pelo site Planejamento Fácil. No Atende, abra a Área do Professor (Diários de Classe) e arraste este arquivo no botão azul: a extensão abre o diário de cada turma e matéria. Não mude os títulos das colunas.'
      + (tipo === 'especial' ? ' O texto entra em Observações, depois do que já estiver no dia.' : ''));
    if (turmas.length === 1) { c += tabela([1800, 5400], [['Turma', turmas[0]]], false); c += par(''); }
    if (tipo === 'especial') {
      c += tabela([1400, 1700, 2300, 2700, 6800], [['Data', 'Turma', 'Disciplina', 'Professor(a)', 'O que o aluno fez (Educação Especial)'],
        ...itens.map(i => [br(i.data), i.turma, i.disciplina, i.professor || '', i.texto])], true);
    } else {
      c += tabela([1300, 1600, 2100, 2500, 4900, 2500], [['Data', 'Turma', 'Disciplina', 'Professor(a)', 'O que foi feito na aula', 'Observações'],
        ...itens.map(i => [br(i.data), i.turma, i.disciplina, i.professor || '', i.descricao, i.observacoes])], true);
    }
    return docx(c, true);
  }

  // ---------- Baixar arquivo (site normal ou dentro do Claude) ----------
  let baixador;
  async function baixar(nome, blob) {
    if (baixador === undefined) {
      baixador = null;
      try { if (window.claude && window.claude.use) baixador = await window.claude.use('downloads'); } catch (e) { }
    }
    if (baixador) {
      try { await baixador.save({ filename: nome, data: blob }); return 'ok'; }
      catch (e) { return e && e.code === 'declined' ? 'cancelado' : 'erro'; }
    }
    const url = URL.createObjectURL(blob), a = document.createElement('a');
    a.href = url; a.download = nome; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    return 'ok';
  }
  const limpaNome = s => String(s || '').replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, ' ').trim();

  // ---------- Abas ----------
  const ABAS = ['inicio', 'planejamento', 'diario', 'especial'];
  function abrirAba(nome) {
    if (!ABAS.includes(nome)) nome = 'inicio';
    for (const a of ABAS) {
      $('#' + a).hidden = a !== nome;
      $('#aba-' + a).setAttribute('aria-selected', String(a === nome));
    }
    if (nome === 'planejamento') desenharPlanejamentos();
    if (nome === 'diario' || nome === 'especial') desenharDiario(nome);
    try { localStorage.setItem(CHAVE + '-aba', nome); } catch (e) { }
  }
  document.querySelectorAll('[data-aba]').forEach(b => b.addEventListener('click', () => { abrirAba(b.dataset.aba); window.scrollTo(0, 0); }));
  document.querySelectorAll('[data-ir]').forEach(b => b.addEventListener('click', () => { abrirAba(b.dataset.ir); window.scrollTo(0, 0); }));

  // ---------- Planejamento ----------
  let editando = null; // id do planejamento aberto
  let confirmandoExcluir = false;

  function desenharPlanejamentos() {
    const sec = $('#planejamento');
    if (editando && dados.planejamentos.some(p => p.id === editando)) return desenharEditor(sec);
    editando = null;
    const lista = [...dados.planejamentos].sort((a, b) => (b.atualizado || 0) - (a.atualizado || 0));
    sec.innerHTML = `
      <div><h2>Planejamento</h2>
        <p class="explica">Um planejamento por turma e disciplina em cada trimestre. Escreva com calma: tudo é guardado sozinho enquanto você digita. Quando terminar, clique em “Gerar Word” e arraste o arquivo no Atende.</p></div>
      <div class="botoes"><button class="btn principal" id="pl-novo">Novo planejamento</button></div>
      <div class="lista">${lista.length ? lista.map(p => `
        <div class="item"><div class="texto"><strong>${esc(p.turma || 'Sem turma')} · ${esc(p.disciplina || 'Sem disciplina')}</strong>
          <span class="meta">${esc(p.trimestre || '')}${p.tema ? ' · ' + esc(p.tema) : ''}${p.inicio ? ' · de ' + br(p.inicio) + ' a ' + br(p.termino) : ''}</span></div>
          <div class="botoes"><button class="btn pequeno" data-abrir="${p.id}">Abrir</button><button class="btn pequeno principal" data-word="${p.id}">Gerar Word</button></div></div>`).join('')
        : '<div class="vazio">Nenhum planejamento ainda. Clique em “Novo planejamento” para começar.</div>'}</div>`;
    $('#pl-novo').onclick = () => {
      const u = dados.ultimo;
      const p = { id: novoId(), turma: u.turma || '', disciplina: u.disciplina || '', trimestre: u.trimestre || '', professor: u.professorPl || '', tema: '', inicio: '', termino: '', atualizado: Date.now() };
      SECOES.forEach(([k]) => p[k] = '');
      dados.planejamentos.push(p); guardar(); editando = p.id; desenharPlanejamentos();
    };
    sec.querySelectorAll('[data-abrir]').forEach(b => b.onclick = () => { editando = b.dataset.abrir; desenharPlanejamentos(); window.scrollTo(0, 0); });
    sec.querySelectorAll('[data-word]').forEach(b => b.onclick = () => gerarPlanejamento(dados.planejamentos.find(p => p.id === b.dataset.word), null));
  }

  function faltandoPlanejamento(p) {
    const f = [];
    if (!p.turma) f.push('Turma'); if (!p.disciplina) f.push('Disciplina'); if (!p.professor) f.push('Professor(a)'); if (!p.trimestre) f.push('Trimestre');
    if (!p.tema) f.push('Tema/Projeto'); if (!p.inicio) f.push('Data de início'); if (!p.termino) f.push('Data de término');
    return f;
  }
  async function gerarPlanejamento(p, msgEl) {
    const f = faltandoPlanejamento(p);
    const nome = limpaNome(`Planejamento ${p.disciplina || ''} ${p.turma || ''} ${p.trimestre || ''}`) + '.docx';
    const r = await baixar(nome, wordPlanejamento(p));
    if (msgEl) {
      msgEl.className = r === 'ok' ? 'aviso ok' : 'aviso';
      msgEl.textContent = r === 'ok' ? `Word gerado: ${nome}. No Atende, abra a Área do Professor na aba Planejamentos e arraste o arquivo no botão azul: a extensão acha o card, abre o Incluir Projeto e preenche. Quem confirma é você.${f.length ? ' Ainda falta preencher: ' + f.join(', ') + '.' : ''}`
        : r === 'cancelado' ? 'Você cancelou o download.' : 'Não consegui gerar o arquivo neste navegador.';
      msgEl.hidden = false;
    }
  }

  function desenharEditor(sec) {
    const p = dados.planejamentos.find(q => q.id === editando);
    const TRI = ['1º Trimestre', '2º Trimestre', '3º Trimestre'];
    sec.innerHTML = `
      <div class="botoes"><button class="btn pequeno" id="pl-voltar">← Meus planejamentos</button><span class="salvo" id="pl-salvo" role="status"></span></div>
      <div><h2>${esc(p.disciplina || 'Novo planejamento')}${p.turma ? ' · ' + esc(p.turma) : ''}</h2>
        <p class="explica">Escreva cada parte na sua caixa. Pode usar Enter para pular linha. Nas caixas de Instrumentos, Critérios, Objetos e Habilidades, coloque um item por linha.</p></div>
      <div class="cabecalho">
        <label class="campo" for="pl-turma">Turma <input id="pl-turma" data-k="turma" placeholder="Ex.: 5º Ano 4" value="${esc(p.turma)}"></label>
        <label class="campo" for="pl-disciplina">Disciplina <input id="pl-disciplina" data-k="disciplina" placeholder="Ex.: Matemática" value="${esc(p.disciplina)}"></label>
        <label class="campo" for="pl-trimestre">Trimestre <select id="pl-trimestre" data-k="trimestre"><option value="">Escolha…</option>${TRI.map(t => `<option${p.trimestre === t ? ' selected' : ''}>${t}</option>`).join('')}</select></label>
        <label class="campo largo" for="pl-professor">Professor(a) <small>Como aparece no card do Atende (ex.: Silvana Strey Miranda). A extensão só abre o planejamento se turma, disciplina e professor baterem com o card.</small>
          <input id="pl-professor" data-k="professor" placeholder="Ex.: Silvana Strey Miranda" value="${esc(p.professor || '')}"></label>
        <label class="campo largo" for="pl-tema">Tema do projeto <input id="pl-tema" data-k="tema" placeholder="Ex.: Frações no dia a dia" value="${esc(p.tema)}"></label>
        <label class="campo" for="pl-inicio">Data de início <input id="pl-inicio" type="date" data-k="inicio" value="${esc(p.inicio)}"></label>
        <label class="campo" for="pl-termino">Data de término <input id="pl-termino" type="date" data-k="termino" value="${esc(p.termino)}"></label>
      </div>
      <div class="caixas">${SECOES.map(([k, nome, d]) => `
        <label class="caixa campo" for="pl-${k}">${esc(nome)}${d ? `<small>${esc(d)}</small>` : ''}
          <textarea id="pl-${k}" data-k="${k}" rows="6">${esc(p[k])}</textarea></label>`).join('')}</div>
      <div class="botoes">
        <button class="btn principal" id="pl-gerar">Gerar Word</button>
        <button class="btn perigo" id="pl-excluir">${confirmandoExcluir ? 'Clique de novo para excluir' : 'Excluir planejamento'}</button>
      </div>
      <p class="aviso" id="pl-msg" role="status" hidden></p>`;
    let t;
    sec.querySelectorAll('[data-k]').forEach(el => el.addEventListener('input', () => {
      p[el.dataset.k] = el.value; p.atualizado = Date.now();
      if (['turma', 'disciplina', 'trimestre'].includes(el.dataset.k)) dados.ultimo[el.dataset.k] = el.value;
      if (el.dataset.k === 'professor') dados.ultimo.professorPl = el.value;
      clearTimeout(t); t = setTimeout(() => { $('#pl-salvo').textContent = guardar() ? 'Guardado neste computador' : 'Não consegui guardar neste navegador'; }, 400);
    }));
    $('#pl-voltar').onclick = () => { editando = null; confirmandoExcluir = false; guardar(); desenharPlanejamentos(); };
    $('#pl-gerar').onclick = () => { guardar(); gerarPlanejamento(p, $('#pl-msg')); };
    $('#pl-excluir').onclick = () => {
      if (!confirmandoExcluir) { confirmandoExcluir = true; $('#pl-excluir').textContent = 'Clique de novo para excluir'; return; }
      dados.planejamentos = dados.planejamentos.filter(q => q.id !== p.id); confirmandoExcluir = false; editando = null; guardar(); desenharPlanejamentos();
    };
  }

  // ---------- Diário (regente e educação especial) ----------
  const semanaAberta = { diario: segunda(hojeISO()), especial: segunda(hojeISO()) };
  const editandoDia = { diario: null, especial: null };

  function desenharDiario(tipo) {
    const sec = $('#' + tipo), lista = dados[tipo], esp = tipo === 'especial';
    const u = dados.ultimo[tipo] || {};
    const ed = editandoDia[tipo] ? lista.find(i => i.id === editandoDia[tipo]) : null;
    const ini = semanaAberta[tipo], fim = somaDias(ini, 6);
    const daSemana = lista.filter(i => i.data >= ini && i.data <= fim).sort((a, b) => a.data.localeCompare(b.data) || (a.criado || 0) - (b.criado || 0));
    const turmas = [...new Set(daSemana.map(i => i.turma || ''))];
    sec.innerHTML = `
      <div><h2>${esp ? 'Diário da Educação Especial' : 'Diário'}</h2>
        <p class="explica">${esp
          ? 'Anote o que o seu aluno fez em cada aula. No fim da semana, gere o Word e arraste no diário da turma no Atende. O seu texto entra em Observações, depois do que a regente escreveu, sem apagar nada.'
          : 'Anote cada aula logo depois de dar, ou tudo de uma vez no fim da semana. Depois gere o Word da semana e arraste no diário da turma no Atende.'}</p></div>
      <form class="anotar" id="${tipo}-form">
        <h3>${ed ? 'Editar anotação' : 'Anotar aula'}</h3>
        <div class="linha3">
          <label class="campo" for="${tipo}-data">Data <input type="date" id="${tipo}-data" required value="${esc(ed ? ed.data : (u.data && u.data >= ini && u.data <= fim ? u.data : hojeISO()))}"></label>
          <label class="campo" for="${tipo}-turma">Turma <input id="${tipo}-turma" required placeholder="Ex.: 5º Ano 4" value="${esc(ed ? ed.turma : u.turma || '')}"></label>
          <label class="campo" for="${tipo}-disciplina">Disciplina <input id="${tipo}-disciplina" required placeholder="Ex.: Matemática" value="${esc(ed ? ed.disciplina : u.disciplina || '')}"></label>
        </div>
        <label class="campo" for="${tipo}-professor">${esp ? 'Professora regente do diário' : 'Professor(a) do diário'} <small>${esp
          ? 'Como aparece no card do Atende (ex.: Silvana Strey Miranda). A extensão só abre o diário se turma, matéria e professor baterem com o card.'
          : 'Seu nome como aparece no card do Atende. A extensão só abre o diário se turma, matéria e professor baterem com o card.'}</small>
          <input id="${tipo}-professor" required placeholder="Ex.: Silvana Strey Miranda" value="${esc(ed ? ed.professor || '' : u.professor || '')}"></label>
        ${esp ? `
        <label class="campo" for="${tipo}-texto">O que o aluno fez nesta aula <small>Escreva como você escreveria no Atende.</small>
          <textarea id="${tipo}-texto" required rows="4">${esc(ed ? ed.texto : '')}</textarea></label>` : `
        <div class="linha2">
          <label class="campo" for="${tipo}-descricao">O que foi feito na aula <textarea id="${tipo}-descricao" required rows="4">${esc(ed ? ed.descricao : '')}</textarea></label>
          <label class="campo" for="${tipo}-observacoes">Observações <small>Comportamento e ocorrências. Pode deixar vazio.</small><textarea id="${tipo}-observacoes" rows="4">${esc(ed ? ed.observacoes : '')}</textarea></label>
        </div>`}
        <div class="botoes"><button class="btn principal" type="submit">${ed ? 'Salvar alteração' : 'Guardar aula'}</button>
          ${ed ? '<button class="btn" type="button" id="' + tipo + '-cancelar">Cancelar</button>' : ''}
          <span class="salvo" id="${tipo}-salvo" role="status"></span></div>
      </form>
      <div class="semana">
        <div class="nav"><button class="btn pequeno" id="${tipo}-ant" aria-label="Semana anterior">←</button>
          <strong>Semana de ${brCurto(ini)} a ${brCurto(somaDias(ini, 4))}</strong>
          <button class="btn pequeno" id="${tipo}-prox" aria-label="Próxima semana">→</button></div>
        <div class="botoes">${daSemana.length ? `<button class="btn principal pequeno" id="${tipo}-gerar">Gerar Word da semana${turmas.length > 1 ? ' (' + turmas.length + ' turmas)' : ''}</button>` : ''}</div>
      </div>
      <p class="aviso" id="${tipo}-msg" role="status" hidden></p>
      ${daSemana.length ? `<div class="tabela-wrap"><table>
        <thead><tr><th>Data</th><th>Turma</th><th>Disciplina</th><th>Professor(a)</th>${esp ? '<th>O que o aluno fez</th>' : '<th>O que foi feito</th><th>Observações</th>'}<th></th></tr></thead>
        <tbody>${daSemana.map(i => `<tr>
          <td class="data">${brCurto(i.data)}<br><span class="salvo">${DIAS[deISO(i.data).getDay()]}</span></td>
          <td>${esc(i.turma)}</td><td>${esc(i.disciplina)}</td><td>${esc(i.professor || '')}</td>
          ${esp ? `<td><div class="quebra">${esc(i.texto)}</div></td>` : `<td><div class="quebra">${esc(i.descricao)}</div></td><td><div class="quebra">${esc(i.observacoes)}</div></td>`}
          <td class="acoes"><button class="btn pequeno" data-editar="${i.id}">Editar</button> <button class="btn pequeno perigo" data-apagar="${i.id}">Apagar</button></td></tr>`).join('')}</tbody></table></div>`
        : '<div class="vazio">Nenhuma aula anotada nesta semana.</div>'}`;

    $(`#${tipo}-form`).addEventListener('submit', ev => {
      ev.preventDefault();
      const v = id => $(`#${tipo}-${id}`).value.trim();
      const item = ed || { id: novoId(), criado: Date.now() };
      item.data = v('data'); item.turma = v('turma'); item.disciplina = v('disciplina'); item.professor = v('professor');
      if (esp) item.texto = v('texto'); else { item.descricao = v('descricao'); item.observacoes = v('observacoes'); }
      if (!item.data) return;
      if (!ed) lista.push(item);
      dados.ultimo[tipo] = { turma: item.turma, disciplina: item.disciplina, professor: item.professor, data: item.data };
      editandoDia[tipo] = null;
      semanaAberta[tipo] = segunda(item.data);
      const ok = guardar();
      desenharDiario(tipo);
      $(`#${tipo}-salvo`).textContent = ok ? `Aula de ${brCurto(item.data)} guardada.` : 'Não consegui guardar neste navegador.';
      const prox = $(`#${tipo}-data`); if (prox) prox.value = somaDias(item.data, deISO(item.data).getDay() === 5 ? 3 : 1);
      $(`#${tipo}-${esp ? 'texto' : 'descricao'}`).focus();
    });
    const canc = $(`#${tipo}-cancelar`); if (canc) canc.onclick = () => { editandoDia[tipo] = null; desenharDiario(tipo); };
    $(`#${tipo}-ant`).onclick = () => { semanaAberta[tipo] = somaDias(ini, -7); desenharDiario(tipo); };
    $(`#${tipo}-prox`).onclick = () => { semanaAberta[tipo] = somaDias(ini, 7); desenharDiario(tipo); };
    sec.querySelectorAll('[data-editar]').forEach(b => b.onclick = () => { editandoDia[tipo] = b.dataset.editar; desenharDiario(tipo); window.scrollTo(0, 0); });
    sec.querySelectorAll('[data-apagar]').forEach(b => b.onclick = () => {
      if (b.dataset.certeza !== '1') { b.dataset.certeza = '1'; b.textContent = 'Tem certeza?'; return; }
      dados[tipo] = dados[tipo].filter(i => i.id !== b.dataset.apagar); guardar(); desenharDiario(tipo);
    });
    const gerar = $(`#${tipo}-gerar`);
    if (gerar) gerar.onclick = async () => {
      const ordem = [...daSemana].sort((a, b) => (a.turma || '').localeCompare(b.turma || '') || (a.disciplina || '').localeCompare(b.disciplina || '') || (a.professor || '').localeCompare(b.professor || '') || a.data.localeCompare(b.data));
      const nome = limpaNome(`Diário ${esp ? 'Educação Especial ' : ''}semana ${brCurto(ini).replace('/', '-')}`) + '.docx';
      const r = await baixar(nome, wordDiario(tipo, ordem));
      const m = $(`#${tipo}-msg`);
      m.className = r === 'ok' ? 'aviso ok' : 'aviso'; m.hidden = false;
      m.textContent = r === 'ok' ? `Word gerado: ${nome}. No Atende, abra a Área do Professor (Diários de Classe) e arraste o arquivo no botão azul. A extensão abre cada turma e matéria e salva os dias.`
        : r === 'cancelado' ? 'Você cancelou o download.' : 'Não consegui gerar o arquivo neste navegador.';
    };
  }

  // ---------- Cópia de segurança ----------
  $('#backup-salvar').onclick = async () => {
    const blob = new Blob([JSON.stringify(dados, null, 1)], { type: 'application/json' });
    const r = await baixar(`Planejamento Fácil - cópia ${br(hojeISO()).replace(/\//g, '-')}.json`, blob);
    $('#backup-msg').textContent = r === 'ok' ? 'Cópia salva. Guarde o arquivo no e-mail ou num pen drive.' : '';
  };
  $('#backup-abrir').onclick = () => $('#backup-arquivo').click();
  $('#backup-arquivo').onchange = async e => {
    const f = e.target.files[0]; if (!f) return;
    try {
      const s = JSON.parse(await f.text());
      if (!s || !Array.isArray(s.planejamentos) || !Array.isArray(s.diario)) throw new Error();
      // junta com o que já tem, sem duplicar
      for (const k of ['planejamentos', 'diario', 'especial']) {
        const ids = new Set(dados[k].map(i => i.id));
        for (const i of s[k] || []) if (!ids.has(i.id)) dados[k].push(i);
      }
      guardar();
      $('#backup-msg').textContent = 'Cópia restaurada.';
      abrirAba(ABAS.find(a => !$('#' + a).hidden));
    } catch (err) { $('#backup-msg').textContent = 'Esse arquivo não é uma cópia do Planejamento Fácil.'; }
    e.target.value = '';
  };

  // Para os testes conferirem o Word gerado
  window.PlanejamentoSite = { wordPlanejamento, wordDiario };

  let inicial = 'inicio';
  try { inicial = localStorage.getItem(CHAVE + '-aba') || 'inicio'; } catch (e) { }
  abrirAba(inicial);
  if (avisoSemGuardar) $('#backup-msg').textContent = 'Este navegador não deixa guardar. O que você escrever some ao fechar.';
})();
