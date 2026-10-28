// Planejamento Fácil: site onde a professora escreve e gera o Word no modelo que a extensão lê.
// Tudo fica no navegador dela (localStorage). Nada é enviado para servidor nenhum.
(() => {
  const CHAVE = 'planejamento-facil-dados';
  // Mesma parte pública da extensão (licenca.js). Vazia = versão de teste, sem pedir chave.
  const CHAVE_PUBLICA = 'XNFrkGiO63POiPKJFwekNHR2TqI8FLrV977l4_0wGZQ.Usdfr6pmn4018Rg18s9T3M5ZydUXUQYKUdRNF4Dh_rg';
  const $ = (s, r = document) => r.querySelector(s);
  const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const novoId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

  // ---------- Dados ----------
  // aulas: o que a professora dá (turma + matéria + professor do diário). registros: 'AAAA-MM-DD|idDaAula' -> texto do dia.
  let dados = { versao: 2, aulas: [], registros: {}, planejamentos: [], ultimo: {} };
  try { const s = JSON.parse(localStorage.getItem(CHAVE)); if (s && typeof s === 'object') dados = { ...dados, ...s }; } catch (e) { }
  let avisoSemGuardar = false;
  function guardar() {
    try { localStorage.setItem(CHAVE, JSON.stringify(dados)); return true; }
    catch (e) { avisoSemGuardar = true; return false; }
  }

  // Acha (ou cria) a aula com essa turma, matéria e professor
  const norm = s => String(s || '').trim().toLowerCase();
  function aulaDe(turma, disciplina, professor, especial) {
    let a = dados.aulas.find(a => norm(a.turma) === norm(turma) && norm(a.disciplina) === norm(disciplina) && norm(a.professor) === norm(professor) && !!a.especial === !!especial);
    if (!a) { a = { id: novoId(), turma: turma || '', disciplina: disciplina || '', professor: professor || '', especial: !!especial, dias: [1, 2, 3, 4, 5] }; dados.aulas.push(a); }
    return a;
  }
  const junta = (a, b) => [a, b].filter(t => t && t.trim()).join('\n');
  // Lista antiga (versão 1: dados.diario e dados.especial) -> aulas + registros
  function importarAntigo(origem) {
    for (const i of origem.diario || []) {
      if (!i.data) continue;
      const a = aulaDe(i.turma, i.disciplina, i.professor, false), k = i.data + '|' + a.id, r = dados.registros[k] || {};
      dados.registros[k] = { descricao: junta(r.descricao, i.descricao), observacoes: junta(r.observacoes, i.observacoes) };
    }
    for (const i of origem.especial || []) {
      if (!i.data) continue;
      const a = aulaDe(i.turma, i.disciplina, i.professor, true), k = i.data + '|' + a.id, r = dados.registros[k] || {};
      dados.registros[k] = { texto: junta(r.texto, i.texto) };
    }
  }
  if (dados.diario || dados.especial) {
    importarAntigo(dados);
    delete dados.diario; delete dados.especial; dados.versao = 2; guardar();
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


  // itens: { data, turma, disciplina, professor, descricao, observacoes }. Educação especial vai com a descrição vazia e o texto em Observações.
  function wordDiario(a, b) {
    // aceita também o jeito antigo: wordDiario('especial' | 'diario', itens)
    const itens = (Array.isArray(a) ? a : b).map(i => i.texto !== undefined && !i.observacoes ? { ...i, descricao: i.descricao || '', observacoes: i.texto, especial: true } : i);
    const turmas = [...new Set(itens.map(i => i.turma || ''))];
    let c = par('DIÁRIO DA SEMANA', { b: true, tam: 20, centro: true, cor: '1F4E79', depois: 4 });
    c += dica('Gerado pelo site Planejamento Fácil. No Atende, abra a Área do Professor (Diários de Classe) e arraste este arquivo no botão azul: a extensão abre o diário de cada turma e matéria. Não mude os títulos das colunas.'
      + (itens.some(i => i.especial) ? ' O texto da Educação Especial entra em Observações, depois do que já estiver no dia.' : ''));
    if (turmas.length === 1) { c += tabela([1800, 5400], [['Turma', turmas[0]]], false); c += par(''); }
    c += tabela([1300, 1600, 2100, 2500, 4900, 2500], [['Data', 'Turma', 'Disciplina', 'Professor(a)', 'O que foi feito na aula', 'Observações'],
      ...itens.map(i => [br(i.data), i.turma, i.disciplina, i.professor || '', i.descricao || '', i.observacoes || ''])], true);
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

  // ---------- Licença (a mesma chave da extensão) ----------
  const b64 = s => { s = s.replace(/-/g, '+').replace(/_/g, '/'); while (s.length % 4) s += '='; return Uint8Array.from(atob(s), c => c.charCodeAt(0)); };
  async function lerChave(texto) {
    const t = String(texto || '').replace(/\s+/g, '');
    const m = t.match(/^PF1-([A-Za-z0-9_-]+)\.([A-Za-z0-9_-]+)$/);
    if (!m) return { ok: false, erro: 'Essa chave está incompleta. Copie de novo, inteira, começando em PF1-.' };
    try {
      const [x, y] = CHAVE_PUBLICA.split('.');
      const pub = await crypto.subtle.importKey('jwk', { kty: 'EC', crv: 'P-256', x, y, ext: true }, { name: 'ECDSA', namedCurve: 'P-256' }, false, ['verify']);
      if (!await crypto.subtle.verify({ name: 'ECDSA', hash: 'SHA-256' }, pub, b64(m[2]), new TextEncoder().encode(m[1]))) return { ok: false, erro: 'Essa chave não é válida.' };
      const d = JSON.parse(new TextDecoder().decode(b64(m[1])));
      if (d.v && d.v < hojeISO()) return { ok: false, erro: `Essa chave venceu em ${br(d.v)}. Fale com quem te vendeu para renovar.` };
      return { ok: true, m: String(d.m), v: d.v || '', chave: t };
    } catch (e) { return { ok: false, erro: 'Essa chave não é válida.' }; }
  }
  const temChrome = () => { try { return !!(window.chrome && chrome.storage && chrome.storage.local); } catch (e) { return false; } };
  async function chaveGuardada() {
    let t = '';
    try { t = localStorage.getItem('pf-licenca') || ''; } catch (e) { }
    if (!t && temChrome()) t = await new Promise(res => { try { chrome.storage.local.get('licenca', o => res((o && o.licenca) || '')); } catch (e) { res(''); } });
    return t;
  }
  function guardarChave(t) {
    try { localStorage.setItem('pf-licenca', t); } catch (e) { }
    if (temChrome()) try { chrome.storage.local.set({ licenca: t }); } catch (e) { }
  }
  function mostrarLicenca(c) {
    $('#lic-info').textContent = c ? `Licença da matrícula ${c.m}${c.v ? ', válida até ' + br(c.v) : ''}` : 'Versão de teste';
  }
  function abrirPortao(erro) {
    $('#app').hidden = true; $('#portao').hidden = false;
    const e = $('#portao-erro'); e.hidden = !erro; e.textContent = erro || '';
    $('#portao-chave').focus();
  }
  $('#portao-form').addEventListener('submit', async ev => {
    ev.preventDefault();
    const c = await lerChave($('#portao-chave').value);
    if (!c.ok) return abrirPortao(c.erro);
    guardarChave(c.chave); mostrarLicenca(c);
    $('#portao').hidden = true; $('#app').hidden = false; $('#portao-chave').value = '';
  });
  $('#lic-trocar').onclick = () => abrirPortao('');

  // ---------- Abas ----------
  const ABAS = ['diario', 'planejamento', 'aulas', 'ajuda'];
  function abrirAba(nome) {
    if (!ABAS.includes(nome)) nome = 'diario';
    for (const a of ABAS) {
      $('#' + a).hidden = a !== nome;
      $('#aba-' + a).setAttribute('aria-selected', String(a === nome));
    }
    if (nome === 'planejamento') desenharPlanejamentos();
    if (nome === 'diario') desenharDiario();
    if (nome === 'aulas') desenharAulas();
    try { localStorage.setItem(CHAVE + '-aba', nome); } catch (e) { }
  }
  document.addEventListener('click', e => {
    const b = e.target.closest('[data-aba], [data-ir]');
    if (b) { abrirAba(b.dataset.aba || b.dataset.ir); window.scrollTo(0, 0); }
  });

  const DIAS_CURTOS = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sáb'];
  const nomeDias = d => d.length === 5 ? 'todos os dias' : d.length ? [...d].sort().map(n => DIAS_CURTOS[n]).join(', ') : 'nenhum dia';

  // ---------- Minhas aulas ----------
  let editandoAula = null;
  function desenharAulas() {
    const sec = $('#aulas');
    const ed = editandoAula ? dados.aulas.find(a => a.id === editandoAula) : null;
    const u = ed || dados.ultimo.aula || dados.aulas[dados.aulas.length - 1] || {};
    const dias = ed ? ed.dias : (u.dias || [1, 2, 3, 4, 5]);
    const lista = [...dados.aulas].sort((a, b) => a.turma.localeCompare(b.turma) || a.disciplina.localeCompare(b.disciplina));
    sec.innerHTML = `
      <div class="titulo-secao"><h2>Minhas aulas</h2>
        <p class="explica">Cadastre uma vez cada turma e matéria que você dá. No Diário, elas aparecem sozinhas no dia certo, uma caixinha para cada.</p></div>
      <form class="cartao formulario" id="aula-form">
        <h3>${ed ? 'Editar aula' : 'Nova aula'}</h3>
        <div class="grade3">
          <label class="campo">Turma <input id="aula-turma" required placeholder="Ex.: 5º Ano 4" value="${esc(u.turma || '')}"></label>
          <label class="campo">Matéria <input id="aula-disciplina" required placeholder="Ex.: Matemática" value="${esc(ed ? ed.disciplina : '')}"></label>
          <label class="campo">Professor(a) do diário <input id="aula-professor" required placeholder="Ex.: Silvana Strey Miranda" value="${esc(u.professor || '')}"></label>
        </div>
        <p class="dica">Escreva turma, matéria e professor como aparecem no card do Atende. A extensão só abre o diário quando os três batem.</p>
        <fieldset class="dias"><legend>Dias em que tem essa aula</legend>
          ${[1, 2, 3, 4, 5].map(n => `<label class="chip"><input type="checkbox" value="${n}"${dias.includes(n) ? ' checked' : ''}><span>${DIAS_CURTOS[n]}</span></label>`).join('')}
        </fieldset>
        <label class="marcar"><input type="checkbox" id="aula-especial"${(ed ? ed.especial : u.especial) ? ' checked' : ''}>
          <span><strong>Sou professora de Educação Especial nesta aula</strong><small>No professor do diário, ponha a regente da turma. O seu texto entra em Observações, depois do que ela escrever.</small></span></label>
        <div class="botoes"><button class="btn principal" type="submit">${ed ? 'Salvar alteração' : 'Adicionar aula'}</button>
          ${ed ? '<button class="btn" type="button" id="aula-cancelar">Cancelar</button>' : ''}
          <span class="salvo" id="aula-salvo" role="status"></span></div>
      </form>
      <div class="lista">${lista.length ? lista.map(a => `
        <div class="item"><div class="texto"><strong>${esc(a.disciplina)} · ${esc(a.turma)}${a.especial ? ' <span class="selo roxo">Ed. Especial</span>' : ''}</strong>
          <span class="meta">${esc(a.professor)} · ${nomeDias(a.dias)}</span></div>
          <div class="botoes"><button class="btn pequeno" data-editar-aula="${a.id}">Editar</button><button class="btn pequeno perigo" data-apagar-aula="${a.id}">Excluir</button></div></div>`).join('')
        : '<div class="vazio">Nenhuma aula ainda. Comece pela primeira turma e matéria do seu horário.</div>'}</div>
      ${lista.length ? '<div class="botoes"><button class="btn principal" data-ir="diario">Ir para o Diário →</button></div>' : ''}`;

    $('#aula-form').addEventListener('submit', ev => {
      ev.preventDefault();
      const v = id => $('#aula-' + id).value.trim();
      const a = ed || { id: novoId() };
      a.turma = v('turma'); a.disciplina = v('disciplina'); a.professor = v('professor'); a.especial = $('#aula-especial').checked;
      a.dias = [...sec.querySelectorAll('.dias input:checked')].map(i => Number(i.value));
      if (!a.dias.length) a.dias = [1, 2, 3, 4, 5];
      if (!ed) dados.aulas.push(a);
      dados.ultimo.aula = { turma: a.turma, professor: a.professor, especial: a.especial, dias: a.dias };
      editandoAula = null;
      const ok = guardar();
      desenharAulas();
      $('#aula-salvo').textContent = ok ? `${a.disciplina} · ${a.turma} guardada. Pode cadastrar a próxima.` : 'Não consegui guardar neste navegador.';
      $('#aula-disciplina').focus();
    });
    const canc = $('#aula-cancelar'); if (canc) canc.onclick = () => { editandoAula = null; desenharAulas(); };
    sec.querySelectorAll('[data-editar-aula]').forEach(b => b.onclick = () => { editandoAula = b.dataset.editarAula; desenharAulas(); window.scrollTo(0, 0); });
    sec.querySelectorAll('[data-apagar-aula]').forEach(b => b.onclick = () => {
      const id = b.dataset.apagarAula;
      const n = Object.keys(dados.registros).filter(k => k.endsWith('|' + id)).length;
      if (b.dataset.certeza !== '1') { b.dataset.certeza = '1'; b.textContent = n ? `Apaga também ${n} anotaç${n > 1 ? 'ões' : 'ão'}. Clique de novo` : 'Clique de novo para excluir'; return; }
      dados.aulas = dados.aulas.filter(a => a.id !== id);
      for (const k of Object.keys(dados.registros)) if (k.endsWith('|' + id)) delete dados.registros[k];
      guardar(); desenharAulas();
    });
  }

  // ---------- Diário: um dia, uma caixinha por aula ----------
  const diaUtil = s => { const w = deISO(s).getDay(); return w === 6 ? somaDias(s, -1) : w === 0 ? somaDias(s, -2) : s; };
  let diaAberto = diaUtil(hojeISO());
  let mostrarTodas = false;
  const chaveReg = (data, a) => data + '|' + a.id;
  const temTexto = r => !!r && Object.values(r).some(t => String(t || '').trim());
  const registro = (data, a) => dados.registros[chaveReg(data, a)];

  function contagemDia(data) {
    const w = deISO(data).getDay();
    const doDia = dados.aulas.filter(a => a.dias.includes(w));
    const feitas = dados.aulas.filter(a => temTexto(registro(data, a))).length;
    return { total: Math.max(doDia.length, feitas), feitas };
  }
  function itensDaSemana(seg) {
    const fim = somaDias(seg, 6), itens = [];
    for (const [k, r] of Object.entries(dados.registros)) {
      const [data, id] = k.split('|');
      if (data < seg || data > fim || !temTexto(r)) continue;
      const a = dados.aulas.find(x => x.id === id); if (!a) continue;
      itens.push(a.especial
        ? { data, turma: a.turma, disciplina: a.disciplina, professor: a.professor, descricao: '', observacoes: (r.texto || '').trim(), especial: true }
        : { data, turma: a.turma, disciplina: a.disciplina, professor: a.professor, descricao: (r.descricao || '').trim(), observacoes: (r.observacoes || '').trim() });
    }
    return itens.sort((a, b) => a.turma.localeCompare(b.turma) || a.disciplina.localeCompare(b.disciplina) || a.professor.localeCompare(b.professor) || a.data.localeCompare(b.data));
  }

  function desenharDiario() {
    const sec = $('#diario');
    if (!dados.aulas.length) {
      sec.innerHTML = `<div class="titulo-secao"><h2>Diário</h2></div>
        <div class="boas-vindas cartao"><h3>Primeiro, cadastre as suas aulas</h3>
          <p>Diga uma vez quais turmas e matérias você dá e em que dias. Depois é só abrir o dia e escrever em cada caixinha.</p>
          <div class="botoes"><button class="btn principal" data-ir="aulas">Cadastrar minhas aulas</button></div></div>`;
      return;
    }
    const dia = diaAberto, w = deISO(dia).getDay(), seg = segunda(dia);
    const doDia = dados.aulas.filter(a => a.dias.includes(w) || temTexto(registro(dia, a)));
    const resto = dados.aulas.filter(a => !doDia.includes(a));
    const vistas = mostrarTodas ? doDia.concat(resto) : doDia;
    const semana = itensDaSemana(seg);
    sec.innerHTML = `
      <div class="dia-topo">
        <div class="nav-dia">
          <button class="btn redondo" id="d-ant" aria-label="Dia anterior">‹</button>
          <div class="dia-nome"><strong>${DIAS[w][0].toUpperCase() + DIAS[w].slice(1)}</strong><span>${br(dia)}</span></div>
          <button class="btn redondo" id="d-prox" aria-label="Próximo dia">›</button>
        </div>
        <div class="botoes"><input type="date" id="d-data" value="${dia}" aria-label="Escolher data"><button class="btn pequeno" id="d-hoje">Hoje</button></div>
      </div>
      <div class="semana-tira" role="list">${[0, 1, 2, 3, 4].map(i => {
        const d = somaDias(seg, i), c = contagemDia(d);
        const estado = c.total && c.feitas >= c.total ? 'cheio' : c.feitas ? 'meio' : '';
        return `<button role="listitem" class="dia-btn ${estado}${d === dia ? ' atual' : ''}" data-dia="${d}" aria-current="${d === dia}">
          <span class="dn">${DIAS_CURTOS[i + 1]}</span><span class="dd">${brCurto(d)}</span><span class="dc">${c.total ? c.feitas + '/' + c.total : '·'}</span></button>`;
      }).join('')}</div>
      <div class="aulas-dia">${vistas.length ? vistas.map(a => cartaoAula(dia, a)).join('') : '<div class="vazio">Nenhuma aula cadastrada para este dia.</div>'}</div>
      ${resto.length ? `<div class="botoes centro"><button class="btn pequeno" id="d-todas">${mostrarTodas ? 'Mostrar só as aulas deste dia' : `Mostrar as outras aulas (${resto.length})`}</button>
        <button class="btn pequeno" data-ir="aulas">Mudar minhas aulas</button></div>` : '<div class="botoes centro"><button class="btn pequeno" data-ir="aulas">Mudar minhas aulas</button></div>'}
      <div class="barra-semana">
        <div><strong>Semana de ${brCurto(seg)} a ${brCurto(somaDias(seg, 4))}</strong><span id="d-resumo">${semana.length ? semana.length + (semana.length > 1 ? ' aulas anotadas' : ' aula anotada') : 'Nada anotado ainda'}</span></div>
        <button class="btn principal" id="d-gerar"${semana.length ? '' : ' disabled'}>Gerar Word da semana</button>
      </div>
      <p class="aviso" id="d-msg" role="status" hidden></p>`;

    const ir = d => { diaAberto = diaUtil(d); desenharDiario(); };
    $('#d-ant').onclick = () => ir(w === 1 ? somaDias(dia, -3) : somaDias(dia, -1));
    $('#d-prox').onclick = () => ir(w === 5 ? somaDias(dia, 3) : somaDias(dia, 1));
    $('#d-hoje').onclick = () => ir(hojeISO());
    $('#d-data').onchange = e => { if (e.target.value) ir(e.target.value); };
    sec.querySelectorAll('[data-dia]').forEach(b => b.onclick = () => ir(b.dataset.dia));
    const todas = $('#d-todas'); if (todas) todas.onclick = () => { mostrarTodas = !mostrarTodas; desenharDiario(); };

    let t;
    sec.querySelectorAll('.aula-card textarea').forEach(el => el.addEventListener('input', () => {
      const card = el.closest('.aula-card'), a = dados.aulas.find(x => x.id === card.dataset.aula), k = chaveReg(dia, a);
      const r = dados.registros[k] || {};
      r[el.dataset.campo] = el.value;
      if (temTexto(r)) dados.registros[k] = r; else delete dados.registros[k];
      card.classList.toggle('feito', temTexto(r));
      card.querySelector('.selo').textContent = temTexto(r) ? 'Anotado' : 'Falta anotar';
      clearTimeout(t); t = setTimeout(() => { if (!guardar()) $('#d-resumo').textContent = 'Não consegui guardar neste navegador'; else atualizarContas(seg, dia); }, 350);
    }));
    $('#d-gerar').onclick = async () => {
      const itens = itensDaSemana(seg); if (!itens.length) return;
      const nome = limpaNome(`Diário semana ${brCurto(seg).replace('/', '-')}`) + '.docx';
      const r = await baixar(nome, wordDiario(itens));
      const m = $('#d-msg');
      m.className = r === 'ok' ? 'aviso ok' : 'aviso'; m.hidden = false;
      m.textContent = r === 'ok' ? `Word gerado: ${nome}. No Atende, abra a Área do Professor (Diários de Classe) e arraste o arquivo no botão azul. A extensão abre cada turma e matéria e preenche os dias.`
        : r === 'cancelado' ? 'Você cancelou o download.' : 'Não consegui gerar o arquivo neste navegador.';
    };
  }
  // Atualiza a tira da semana e o resumo sem redesenhar as caixinhas (para não perder o cursor)
  function atualizarContas(seg, dia) {
    document.querySelectorAll('#diario [data-dia]').forEach(b => {
      const c = contagemDia(b.dataset.dia);
      b.querySelector('.dc').textContent = c.total ? c.feitas + '/' + c.total : '·';
      b.classList.toggle('cheio', !!c.total && c.feitas >= c.total);
      b.classList.toggle('meio', !!c.feitas && c.feitas < c.total);
    });
    const n = itensDaSemana(seg).length;
    $('#d-resumo').textContent = n ? n + (n > 1 ? ' aulas anotadas' : ' aula anotada') : 'Nada anotado ainda';
    $('#d-gerar').disabled = !n;
  }
  function cartaoAula(dia, a) {
    const r = registro(dia, a) || {}, ok = temTexto(r), id = 'c-' + a.id;
    return `<article class="aula-card${ok ? ' feito' : ''}${a.especial ? ' especial' : ''}" data-aula="${a.id}">
      <header><div><h3>${esc(a.disciplina)}</h3><span class="meta">${esc(a.turma)} · ${esc(a.professor)}${a.especial ? ' · Educação Especial' : ''}</span></div>
        <span class="selo">${ok ? 'Anotado' : 'Falta anotar'}</span></header>
      ${a.especial
        ? `<label class="campo" for="${id}-t">O que o aluno fez<textarea id="${id}-t" data-campo="texto" rows="3">${esc(r.texto || '')}</textarea></label>`
        : `<label class="campo" for="${id}-d">O que foi feito na aula<textarea id="${id}-d" data-campo="descricao" rows="3">${esc(r.descricao || '')}</textarea></label>
           <label class="campo obs" for="${id}-o">Observações <small>opcional</small><textarea id="${id}-o" data-campo="observacoes" rows="1">${esc(r.observacoes || '')}</textarea></label>`}
    </article>`;
  }

  // ---------- Planejamento ----------
  let editando = null; // id do planejamento aberto
  let confirmandoExcluir = false;

  function desenharPlanejamentos() {
    const sec = $('#planejamento');
    if (editando && dados.planejamentos.some(p => p.id === editando)) return desenharEditor(sec);
    editando = null;
    const lista = [...dados.planejamentos].sort((a, b) => (b.atualizado || 0) - (a.atualizado || 0));
    sec.innerHTML = `
      <div class="titulo-secao"><h2>Planejamento</h2>
        <p class="explica">Um planejamento por turma e matéria em cada trimestre. Tudo é guardado sozinho enquanto você digita. Quando terminar, gere o Word e arraste no Atende.</p></div>
      <div class="botoes"><button class="btn principal" id="pl-novo">Novo planejamento</button></div>
      <div class="lista">${lista.length ? lista.map(p => `
        <div class="item"><div class="texto"><strong>${esc(p.disciplina || 'Sem matéria')} · ${esc(p.turma || 'Sem turma')}</strong>
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
    const regulares = dados.aulas.filter(a => !a.especial);
    sec.innerHTML = `
      <div class="botoes"><button class="btn pequeno" id="pl-voltar">← Meus planejamentos</button><span class="salvo" id="pl-salvo" role="status"></span></div>
      <div class="titulo-secao"><h2>${esc(p.disciplina || 'Novo planejamento')}${p.turma ? ' · ' + esc(p.turma) : ''}</h2>
        <p class="explica">Escreva cada parte na sua caixa. Nas caixas de Instrumentos, Critérios, Objetos e Habilidades, coloque um item por linha.</p></div>
      <div class="cartao cabecalho">
        ${regulares.length ? `<label class="campo largo" for="pl-puxar">Puxar de Minhas aulas <select id="pl-puxar"><option value="">Escolha uma aula para preencher turma, matéria e professor…</option>
          ${regulares.map(a => `<option value="${a.id}">${esc(a.disciplina)} · ${esc(a.turma)} · ${esc(a.professor)}</option>`).join('')}</select></label>` : ''}
        <label class="campo" for="pl-turma">Turma <input id="pl-turma" data-k="turma" placeholder="Ex.: 5º Ano 4" value="${esc(p.turma)}"></label>
        <label class="campo" for="pl-disciplina">Matéria <input id="pl-disciplina" data-k="disciplina" placeholder="Ex.: Matemática" value="${esc(p.disciplina)}"></label>
        <label class="campo" for="pl-trimestre">Trimestre <select id="pl-trimestre" data-k="trimestre"><option value="">Escolha…</option>${TRI.map(t => `<option${p.trimestre === t ? ' selected' : ''}>${t}</option>`).join('')}</select></label>
        <label class="campo largo" for="pl-professor">Professor(a) <small>Como aparece no card do Atende. A extensão só abre o planejamento se turma, matéria e professor baterem com o card.</small>
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
    const mudou = (k, v) => {
      p[k] = v; p.atualizado = Date.now();
      if (['turma', 'disciplina', 'trimestre'].includes(k)) dados.ultimo[k] = v;
      if (k === 'professor') dados.ultimo.professorPl = v;
      clearTimeout(t); t = setTimeout(() => { $('#pl-salvo').textContent = guardar() ? 'Guardado neste computador' : 'Não consegui guardar neste navegador'; }, 400);
    };
    sec.querySelectorAll('[data-k]').forEach(el => el.addEventListener('input', () => mudou(el.dataset.k, el.value)));
    const puxar = $('#pl-puxar');
    if (puxar) puxar.onchange = () => {
      const a = dados.aulas.find(x => x.id === puxar.value); if (!a) return;
      for (const k of ['turma', 'disciplina', 'professor']) { $('#pl-' + k).value = a[k]; mudou(k, a[k]); }
      sec.querySelector('h2').textContent = a.disciplina + ' · ' + a.turma;
    };
    $('#pl-voltar').onclick = () => { editando = null; confirmandoExcluir = false; guardar(); desenharPlanejamentos(); };
    $('#pl-gerar').onclick = () => { guardar(); gerarPlanejamento(p, $('#pl-msg')); };
    $('#pl-excluir').onclick = () => {
      if (!confirmandoExcluir) { confirmandoExcluir = true; $('#pl-excluir').textContent = 'Clique de novo para excluir'; return; }
      dados.planejamentos = dados.planejamentos.filter(q => q.id !== p.id); confirmandoExcluir = false; editando = null; guardar(); desenharPlanejamentos();
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
      if (!s || !Array.isArray(s.planejamentos) || !(Array.isArray(s.diario) || Array.isArray(s.aulas))) throw new Error();
      // junta com o que já tem, sem duplicar nem trocar o que já está aqui
      const ids = new Set(dados.planejamentos.map(i => i.id));
      for (const i of s.planejamentos) if (!ids.has(i.id)) dados.planejamentos.push(i);
      if (Array.isArray(s.aulas)) {
        const aIds = new Set(dados.aulas.map(a => a.id));
        for (const a of s.aulas) if (!aIds.has(a.id)) dados.aulas.push(a);
        for (const [k, r] of Object.entries(s.registros || {})) if (!dados.registros[k]) dados.registros[k] = r;
      } else importarAntigo(s);
      guardar();
      $('#backup-msg').textContent = 'Cópia restaurada.';
      abrirAba(ABAS.find(a => !$('#' + a).hidden));
    } catch (err) { $('#backup-msg').textContent = 'Esse arquivo não é uma cópia do Planejamento Fácil.'; }
    e.target.value = '';
  };

  // Para os testes conferirem o Word gerado
  window.PlanejamentoSite = { wordPlanejamento, wordDiario };

  let inicial = 'diario';
  try { inicial = localStorage.getItem(CHAVE + '-aba') || 'diario'; } catch (e) { }
  if (!dados.aulas.length && !dados.planejamentos.length && inicial === 'diario') inicial = 'diario';
  abrirAba(inicial);
  if (avisoSemGuardar) $('#backup-msg').textContent = 'Este navegador não deixa guardar. O que você escrever some ao fechar.';

  (async () => {
    if (!CHAVE_PUBLICA) { mostrarLicenca(null); $('#app').hidden = false; return; }
    const t = await chaveGuardada();
    const c = t ? await lerChave(t) : { ok: false, erro: '' };
    if (c.ok) { mostrarLicenca(c); $('#app').hidden = false; }
    else abrirPortao(c.erro);
  })();
})();
