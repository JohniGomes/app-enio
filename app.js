'use strict';

// ================================================================
//  ERGO X — app.js
//  Cole a URL do seu Apps Script abaixo e remova o aviso de setup.
// ================================================================

const CONFIG = {
  API_URL: 'https://script.google.com/macros/s/AKfycbzkkG-HAbqLGbInZYH2zPE02QAw0czwXZ6QT0b_sqMVBSv1CreAwv4aYmXvIpgQryah/exec',
  SHEETS: { AET: 'AET', PA: 'PA' }
};

// ================================================================
//  STATE
// ================================================================
const State = {
  aet: [],
  pa:  [],
  charts: { aetGenero: null, aetCrit: null, paStatus: null, paCrit: null },
  editTarget: null
};

// ================================================================
//  API — todas as operações via GET para evitar CORS com Apps Script
// ================================================================
const API = {
  isConfigured() {
    return CONFIG.API_URL && CONFIG.API_URL !== 'https://script.google.com/macros/s/AKfycbzkkG-HAbqLGbInZYH2zPE02QAw0czwXZ6QT0b_sqMVBSv1CreAwv4aYmXvIpgQryah/exec';
  },

  encodeData(obj) {
    const json = JSON.stringify(obj);
    const bytes = new TextEncoder().encode(json);
    const bin   = Array.from(bytes).map(b => String.fromCharCode(b)).join('');
    return btoa(bin);
  },

  async call(params) {
    const url = new URL(CONFIG.API_URL);
    Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, String(v)));
    const res  = await fetch(url.toString());
    const json = await res.json();
    if (!json.success) throw new Error(json.error || 'Erro desconhecido');
    return json.data;
  },

  read(sheet)                  { return this.call({ action: 'read', sheet }); },
  create(sheet, data)          { return this.call({ action: 'create', sheet, data: this.encodeData(data) }); },
  update(sheet, rowNum, data)  { return this.call({ action: 'update', sheet, rowNum, data: this.encodeData(data) }); },
  delete(sheet, rowNum)        { return this.call({ action: 'delete', sheet, rowNum }); }
};

// ================================================================
//  UTILS
// ================================================================
const Utils = {
  esc(str) {
    return String(str || '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  },

  truncate(str, n = 75) {
    const s = str || '';
    return s.length > n ? s.slice(0, n) + '…' : s;
  },

  formatDate(val) {
    if (!val) return '—';
    const d = new Date(val);
    return isNaN(d) ? String(val) : d.toLocaleDateString('pt-BR');
  },

  unique(arr, key) {
    return [...new Set(arr.map(r => r[key]).filter(Boolean))].sort();
  },

  fillSelect(id, values) {
    const sel = document.getElementById(id);
    if (!sel) return;
    const cur   = sel.value;
    const first = sel.options[0].cloneNode(true);
    sel.innerHTML = '';
    sel.appendChild(first);
    values.forEach(v => {
      const o = document.createElement('option');
      o.value = o.textContent = v;
      sel.appendChild(o);
    });
    if (cur) sel.value = cur;
  },

  fillDatalist(id, values) {
    const dl = document.getElementById(id);
    if (dl) dl.innerHTML = values.map(v => `<option value="${Utils.esc(v)}">`).join('');
  },

  critBadge(val) {
    const map = { 'ALTO': 'alto', 'MODERADO': 'moderado', 'BAIXO': 'baixo',
                  'AUSÊNCIA DE RISCO': 'ausencia', 'EXTINTO': 'extinto', 'DESATIVADO': 'extinto' };
    const cls = map[(val || '').toUpperCase()] || 'extinto';
    return `<span class="badge badge-${cls}">${Utils.esc(val) || '—'}</span>`;
  },

  semaforo(r) {
    if (r.STATUS === 'OK' || r.DATA_CONCLUSAO) return { label: 'CONCLUÍDO', cls: 'sem-verde' };
    if (r.DATA_PREVISTA && new Date(r.DATA_PREVISTA) < new Date()) return { label: 'ATRASADO', cls: 'sem-vermelho' };
    return { label: 'EM ANDAMENTO', cls: 'sem-amarelo' };
  },

  toast(msg, type = '') {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.className   = `toast show ${type}`;
    clearTimeout(el._t);
    el._t = setTimeout(() => (el.className = 'toast'), 3500);
  },

  critOpts(selected) {
    return ['ALTO','MODERADO','BAIXO','AUSÊNCIA DE RISCO','EXTINTO']
      .map(c => `<option${c === selected ? ' selected' : ''}>${c}</option>`).join('');
  },

  critPaOpts(selected) {
    return ['ALTO','MODERADO','BAIXO']
      .map(c => `<option${c === selected ? ' selected' : ''}>${c}</option>`).join('');
  },

  selectOpts(list, selected) {
    return list.map(v => `<option value="${v}"${v === selected ? ' selected' : ''}>${v}</option>`).join('');
  },

  dateValue(val) {
    if (!val) return '';
    const d = new Date(val);
    return isNaN(d) ? '' : d.toISOString().split('T')[0];
  }
};

// ================================================================
//  CHARTS helpers
// ================================================================
const Charts = {
  destroy(key) {
    if (State.charts[key]) { State.charts[key].destroy(); State.charts[key] = null; }
  },

  donut(key, canvasId, labels, data, colors) {
    this.destroy(key);
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;
    State.charts[key] = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels,
        datasets: [{ data, backgroundColor: colors, borderWidth: 2, borderColor: '#fff' }]
      },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '65%',
        plugins: {
          legend: { position: 'bottom', labels: { boxWidth: 11, padding: 10, font: { size: 11 } } }
        }
      }
    });
  },

  bar(key, canvasId, labels, datasets, stacked = false) {
    this.destroy(key);
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;
    State.charts[key] = new Chart(ctx, {
      type: 'bar',
      data: { labels, datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        scales: {
          x: { stacked, ticks: { font: { size: 10 }, maxRotation: 35 } },
          y: { stacked, beginAtZero: true, ticks: { precision: 0 } }
        },
        plugins: {
          legend: {
            display: datasets.length > 1,
            position: 'bottom',
            labels: { boxWidth: 11, padding: 10, font: { size: 11 } }
          }
        }
      }
    });
  }
};

// ================================================================
//  AET MODULE
// ================================================================
const AET = {
  f: { setor: '', criticidade: '', genero: '', gerente: '', q: '' },

  load(data) {
    State.aet = data;
    Utils.fillSelect('aet-filter-setor',   Utils.unique(data, 'SETOR'));
    Utils.fillSelect('aet-filter-gerente', Utils.unique(data, 'GERENTE'));
    Utils.fillDatalist('dl-setor',   Utils.unique(data, 'SETOR'));
    Utils.fillDatalist('dl-gerente', Utils.unique(data, 'GERENTE'));
    Utils.fillDatalist('dl-setor-pa',   Utils.unique(data, 'SETOR'));
    Utils.fillDatalist('dl-gerente-pa', Utils.unique(data, 'GERENTE'));
    this.apply();
  },

  onFilter() {
    this.f.setor       = document.getElementById('aet-filter-setor').value;
    this.f.criticidade = document.getElementById('aet-filter-criticidade').value;
    this.f.genero      = document.getElementById('aet-filter-genero').value;
    this.f.gerente     = document.getElementById('aet-filter-gerente').value;
    this.apply();
  },

  search(q) { this.f.q = q; this.apply(); },

  clearFilters() {
    this.f = { setor: '', criticidade: '', genero: '', gerente: '', q: '' };
    ['aet-filter-setor','aet-filter-criticidade','aet-filter-genero','aet-filter-gerente']
      .forEach(id => document.getElementById(id).value = '');
    document.getElementById('aet-search').value = '';
    this.apply();
  },

  apply() {
    let d = [...State.aet];
    const f = this.f;
    if (f.setor)       d = d.filter(r => r.SETOR === f.setor);
    if (f.criticidade) d = d.filter(r => (r.CRITICIDADE_ATUAL || '').toUpperCase() === f.criticidade.toUpperCase());
    if (f.genero)      d = d.filter(r => r.POSTO_GENERO === f.genero);
    if (f.gerente)     d = d.filter(r => r.GERENTE === f.gerente);
    if (f.q) {
      const q = f.q.toLowerCase();
      d = d.filter(r =>
        (r.SETOR || '').toLowerCase().includes(q) ||
        (r.POSTO_TRABALHO || '').toLowerCase().includes(q) ||
        (r.GERENTE || '').toLowerCase().includes(q)
      );
    }
    this.renderCards(d);
    this.renderCharts(d);
    this.renderTable(d);
  },

  cnt(data, crit) {
    return data.filter(r => (r.CRITICIDADE_ATUAL || '').toUpperCase() === crit.toUpperCase()).length;
  },

  renderCards(d) {
    document.getElementById('aet-total').textContent    = d.length;
    document.getElementById('aet-alto').textContent     = this.cnt(d, 'ALTO');
    document.getElementById('aet-moderado').textContent = this.cnt(d, 'MODERADO');
    document.getElementById('aet-baixo').textContent    = this.cnt(d, 'BAIXO');
    document.getElementById('aet-ausencia').textContent = this.cnt(d, 'AUSÊNCIA DE RISCO');
    document.getElementById('aet-extinto').textContent  = this.cnt(d, 'EXTINTO') + this.cnt(d, 'DESATIVADO');
  },

  renderCharts(d) {
    // Donut gênero
    const gens = ['Masculino', 'Feminino', 'Unissex', '?'];
    Charts.donut('aetGenero', 'aet-chart-genero',
      ['Masculino', 'Feminino', 'Unissex', 'Indefinido'],
      gens.map(g => d.filter(r => r.POSTO_GENERO === g).length),
      ['#1B4472', '#17B3CC', '#82C341', '#95A5A6']
    );

    // Stacked bar por setor
    const setores = Utils.unique(d, 'SETOR').slice(0, 12);
    const crits   = ['ALTO', 'MODERADO', 'BAIXO'];
    const colors  = { ALTO: '#E74C3C', MODERADO: '#F39C12', BAIXO: '#27AE60' };
    Charts.bar('aetCrit', 'aet-chart-crit', setores,
      crits.map(c => ({
        label: c.charAt(0) + c.slice(1).toLowerCase(),
        data: setores.map(s => d.filter(r => r.SETOR === s && (r.CRITICIDADE_ATUAL || '').toUpperCase() === c).length),
        backgroundColor: colors[c],
        borderRadius: 3
      })), true
    );
  },

  renderTable(d) {
    const tbody = document.getElementById('aet-tbody');
    document.getElementById('aet-count').textContent = `${d.length} registro${d.length !== 1 ? 's' : ''}`;
    if (!d.length) {
      tbody.innerHTML = '<tr><td colspan="8" class="table-loading">Nenhum registro encontrado.</td></tr>';
      return;
    }
    tbody.innerHTML = d.map(r => `
      <tr>
        <td>${Utils.esc(r.ID) || '—'}</td>
        <td>${Utils.esc(r.SETOR) || '—'}</td>
        <td title="${Utils.esc(r.POSTO_TRABALHO)}">${Utils.esc(Utils.truncate(r.POSTO_TRABALHO, 50))}</td>
        <td>${Utils.critBadge(r.CRITICIDADE_ATUAL)}</td>
        <td>${Utils.esc(r.POSTO_GENERO) || '—'}</td>
        <td>${Utils.esc(r.GERENTE) || '—'}</td>
        <td>${Utils.esc(r.ATUALIZACAO) || '—'}</td>
        <td>
          <div class="action-group">
            <button class="btn-action btn-edit" onclick="Modal.openAET(${r._row})">Editar</button>
            <button class="btn-action btn-delete" onclick="AET.confirmDelete(${r._row})">Excluir</button>
          </div>
        </td>
      </tr>`).join('');
  },

  async confirmDelete(rowNum) {
    const r = State.aet.find(x => x._row === rowNum);
    if (!r) return;
    if (!confirm(`Excluir o posto:\n"${r.POSTO_TRABALHO}"\n\nEsta ação não pode ser desfeita.`)) return;
    try {
      await API.delete(CONFIG.SHEETS.AET, rowNum);
      Utils.toast('Posto excluído com sucesso.', 'success');
      await App.loadAET();
    } catch (e) {
      Utils.toast('Erro ao excluir: ' + e.message, 'error');
    }
  }
};

// ================================================================
//  PA MODULE
// ================================================================
const PA = {
  f: { setor: '', criticidade: '', status: '', gerente: '', q: '' },

  getSem(r) { return Utils.semaforo(r); },

  load(data) {
    State.pa = data.map(r => ({ ...r, _sem: this.getSem(r) }));
    Utils.fillSelect('pa-filter-setor',   Utils.unique(data, 'SETOR'));
    Utils.fillSelect('pa-filter-gerente', Utils.unique(data, 'GERENTE'));
    this.apply();
  },

  onFilter() {
    this.f.setor       = document.getElementById('pa-filter-setor').value;
    this.f.criticidade = document.getElementById('pa-filter-criticidade').value;
    this.f.status      = document.getElementById('pa-filter-status').value;
    this.f.gerente     = document.getElementById('pa-filter-gerente').value;
    this.apply();
  },

  search(q) { this.f.q = q; this.apply(); },

  clearFilters() {
    this.f = { setor: '', criticidade: '', status: '', gerente: '', q: '' };
    ['pa-filter-setor','pa-filter-criticidade','pa-filter-status','pa-filter-gerente']
      .forEach(id => document.getElementById(id).value = '');
    document.getElementById('pa-search').value = '';
    this.apply();
  },

  apply() {
    let d = [...State.pa];
    const f = this.f;
    if (f.setor)       d = d.filter(r => r.SETOR === f.setor);
    if (f.criticidade) d = d.filter(r => (r.CRITICIDADE || '').toUpperCase() === f.criticidade.toUpperCase());
    if (f.status)      d = d.filter(r => r._sem.label === f.status.toUpperCase());
    if (f.gerente)     d = d.filter(r => r.GERENTE === f.gerente);
    if (f.q) {
      const q = f.q.toLowerCase();
      d = d.filter(r =>
        (r.SETOR || '').toLowerCase().includes(q) ||
        (r.POSTO_TRABALHO || '').toLowerCase().includes(q) ||
        (r.ACAO_CONTROLE || '').toLowerCase().includes(q) ||
        (r.RESPONSAVEL || '').toLowerCase().includes(q)
      );
    }
    this.renderCards(d);
    this.renderCharts(d);
    this.renderTable(d);
  },

  renderCards(d) {
    const c = d.filter(r => r._sem.label === 'CONCLUÍDO').length;
    const a = d.filter(r => r._sem.label === 'ATRASADO').length;
    const e = d.filter(r => r._sem.label === 'EM ANDAMENTO').length;
    document.getElementById('pa-total').textContent    = d.length;
    document.getElementById('pa-concluido').textContent = c;
    document.getElementById('pa-atrasado').textContent  = a;
    document.getElementById('pa-andamento').textContent = e;
    document.getElementById('pa-pct').textContent = d.length ? Math.round(c / d.length * 100) + '%' : '—';
  },

  renderCharts(d) {
    const c = d.filter(r => r._sem.label === 'CONCLUÍDO').length;
    const a = d.filter(r => r._sem.label === 'ATRASADO').length;
    const e = d.filter(r => r._sem.label === 'EM ANDAMENTO').length;
    Charts.donut('paStatus', 'pa-chart-status',
      ['Concluído', 'Atrasado', 'Em Andamento'], [c, a, e],
      ['#27AE60', '#E74C3C', '#F39C12']
    );
    Charts.bar('paCrit', 'pa-chart-crit',
      ['Alto', 'Moderado', 'Baixo'],
      [{ label: 'Ações',
         data: ['ALTO','MODERADO','BAIXO'].map(x => d.filter(r => (r.CRITICIDADE||'').toUpperCase() === x).length),
         backgroundColor: ['#E74C3C','#F39C12','#27AE60'],
         borderRadius: 5
      }]
    );
  },

  renderTable(d) {
    const tbody = document.getElementById('pa-tbody');
    document.getElementById('pa-count').textContent = `${d.length} registro${d.length !== 1 ? 's' : ''}`;
    if (!d.length) {
      tbody.innerHTML = '<tr><td colspan="8" class="table-loading">Nenhum registro encontrado.</td></tr>';
      return;
    }
    tbody.innerHTML = d.map(r => `
      <tr>
        <td>${Utils.esc(r.SETOR) || '—'}</td>
        <td title="${Utils.esc(r.POSTO_TRABALHO)}">${Utils.esc(Utils.truncate(r.POSTO_TRABALHO, 35))}</td>
        <td>${Utils.critBadge(r.CRITICIDADE)}</td>
        <td title="${Utils.esc(r.ACAO_CONTROLE)}">${Utils.esc(Utils.truncate(r.ACAO_CONTROLE, 70))}</td>
        <td>${Utils.esc(r.RESPONSAVEL || r.GERENTE) || '—'}</td>
        <td>${Utils.formatDate(r.DATA_PREVISTA)}</td>
        <td><span class="semaforo ${r._sem.cls}">${r._sem.label}</span></td>
        <td>
          <div class="action-group">
            <button class="btn-action btn-edit" onclick="Modal.openPA(${r._row})">Editar</button>
            <button class="btn-action btn-delete" onclick="PA.confirmDelete(${r._row})">Excluir</button>
          </div>
        </td>
      </tr>`).join('');
  },

  async confirmDelete(rowNum) {
    if (!confirm('Excluir esta ação de controle?\n\nEsta ação não pode ser desfeita.')) return;
    try {
      await API.delete(CONFIG.SHEETS.PA, rowNum);
      Utils.toast('Ação excluída com sucesso.', 'success');
      await App.loadPA();
    } catch (e) {
      Utils.toast('Erro ao excluir: ' + e.message, 'error');
    }
  }
};

// ================================================================
//  MODAL
// ================================================================
const Modal = {
  open(title, html) {
    document.getElementById('modal-title').textContent = title;
    document.getElementById('modal-body').innerHTML    = html;
    document.getElementById('modal').classList.add('open');
    document.body.style.overflow = 'hidden';
  },

  close() {
    document.getElementById('modal').classList.remove('open');
    document.body.style.overflow = '';
    State.editTarget = null;
  },

  closeOnOverlay(e) {
    if (e.target === document.getElementById('modal')) this.close();
  },

  openAET(rowNum) {
    const r = State.aet.find(x => x._row === rowNum);
    if (!r) return;
    State.editTarget = { sheet: 'AET', rowNum, orig: r };

    const yrs = ['2024','2023','2022','2021','2020','2019'];
    const yFields = yrs.map(y => {
      const key = `CRITICIDADE_${y}`;
      const v   = r[key] || '';
      const opts = ['','ALTO','MODERADO','BAIXO','AUSÊNCIA DE RISCO']
        .map(c => `<option value="${c}"${c === v ? ' selected' : ''}>${c || '—'}</option>`).join('');
      return `<div class="form-group"><label>Criticidade ${y}</label><select name="${key}">${opts}</select></div>`;
    }).join('');

    const genOpts = ['','Masculino','Feminino','Unissex','?']
      .map(g => `<option value="${g}"${g === (r.POSTO_GENERO||'') ? ' selected' : ''}>${g || 'Selecione…'}</option>`).join('');

    this.open('Editar Posto de Trabalho', `
      <div class="form-grid" id="modal-form-aet">
        <div class="form-group">
          <label>Setor *</label>
          <input type="text" id="me-SETOR" value="${Utils.esc(r.SETOR)}" required>
        </div>
        <div class="form-group">
          <label>Posto de Trabalho *</label>
          <input type="text" id="me-POSTO_TRABALHO" value="${Utils.esc(r.POSTO_TRABALHO)}" required>
        </div>
        <div class="form-group">
          <label>Criticidade Atual *</label>
          <select id="me-CRITICIDADE_ATUAL"><option value="">Selecione…</option>${Utils.critOpts(r.CRITICIDADE_ATUAL)}</select>
        </div>
        <div class="form-group">
          <label>Gênero do Posto</label>
          <select id="me-POSTO_GENERO">${genOpts}</select>
        </div>
        <div class="form-group">
          <label>Gerente</label>
          <input type="text" id="me-GERENTE" value="${Utils.esc(r.GERENTE)}">
        </div>
        <div class="form-group">
          <label>Atualização</label>
          <input type="text" id="me-ATUALIZACAO" value="${Utils.esc(r.ATUALIZACAO)}">
        </div>
        ${yFields.replace(/id="me-/g, 'data-key="').replace(/<select data-key="([^"]+)">/g, (_, k) => `<select id="me-${k}">`)}
        <div class="form-group form-full">
          <label>Observações</label>
          <textarea id="me-OBSERVACOES" rows="3">${Utils.esc(r.OBSERVACOES)}</textarea>
        </div>
        <div class="form-group form-full">
          <label>Condição para Unissex</label>
          <textarea id="me-CONDICAO_UNISSEX" rows="2">${Utils.esc(r.CONDICAO_UNISSEX)}</textarea>
        </div>
        <div class="form-actions form-full">
          <button type="button" class="btn-secondary" onclick="Modal.close()">Cancelar</button>
          <button type="button" class="btn-primary" onclick="Modal.saveAET()">Salvar</button>
        </div>
      </div>`);
  },

  async saveAET() {
    const get = id => (document.getElementById('me-' + id) || {}).value || '';
    const data = {
      ...State.editTarget.orig,
      SETOR:             get('SETOR'),
      POSTO_TRABALHO:    get('POSTO_TRABALHO'),
      CRITICIDADE_ATUAL: get('CRITICIDADE_ATUAL'),
      POSTO_GENERO:      get('POSTO_GENERO'),
      GERENTE:           get('GERENTE'),
      ATUALIZACAO:       get('ATUALIZACAO'),
      CRITICIDADE_2024:  get('CRITICIDADE_2024'),
      CRITICIDADE_2023:  get('CRITICIDADE_2023'),
      CRITICIDADE_2022:  get('CRITICIDADE_2022'),
      CRITICIDADE_2021:  get('CRITICIDADE_2021'),
      CRITICIDADE_2020:  get('CRITICIDADE_2020'),
      CRITICIDADE_2019:  get('CRITICIDADE_2019'),
      OBSERVACOES:       get('OBSERVACOES'),
      CONDICAO_UNISSEX:  get('CONDICAO_UNISSEX')
    };
    if (!data.SETOR || !data.POSTO_TRABALHO) {
      Utils.toast('Preencha os campos obrigatórios.', 'error'); return;
    }
    try {
      await API.update(CONFIG.SHEETS.AET, State.editTarget.rowNum, data);
      Utils.toast('Posto atualizado com sucesso!', 'success');
      this.close();
      await App.loadAET();
    } catch (e) {
      Utils.toast('Erro ao salvar: ' + e.message, 'error');
    }
  },

  openPA(rowNum) {
    const r = State.pa.find(x => x._row === rowNum);
    if (!r) return;
    State.editTarget = { sheet: 'PA', rowNum, orig: r };

    const classOpts = ['', 'Ação Normativa', 'Sugestão de Melhoria', 'Engenharia']
      .map(v => `<option value="${v}"${v === (r.CLASSIFICACAO||'') ? ' selected' : ''}>${v || 'Selecione…'}</option>`).join('');

    this.open('Editar Ação de Controle', `
      <div class="form-grid">
        <div class="form-group">
          <label>Setor *</label>
          <input type="text" id="me-SETOR" value="${Utils.esc(r.SETOR)}" required>
        </div>
        <div class="form-group">
          <label>Posto de Trabalho</label>
          <input type="text" id="me-POSTO_TRABALHO" value="${Utils.esc(r.POSTO_TRABALHO)}">
        </div>
        <div class="form-group">
          <label>Criticidade</label>
          <select id="me-CRITICIDADE"><option value="">Selecione…</option>${Utils.critPaOpts(r.CRITICIDADE)}</select>
        </div>
        <div class="form-group">
          <label>Classificação</label>
          <select id="me-CLASSIFICACAO">${classOpts}</select>
        </div>
        <div class="form-group">
          <label>Gerente</label>
          <input type="text" id="me-GERENTE" value="${Utils.esc(r.GERENTE)}">
        </div>
        <div class="form-group">
          <label>Responsável</label>
          <input type="text" id="me-RESPONSAVEL" value="${Utils.esc(r.RESPONSAVEL)}">
        </div>
        <div class="form-group">
          <label>Estimativa de Valor (R$)</label>
          <input type="number" id="me-ESTIMATIVA_VALOR" value="${Utils.esc(r.ESTIMATIVA_VALOR)}" step="0.01" min="0">
        </div>
        <div class="form-group">
          <label>Data Prevista</label>
          <input type="date" id="me-DATA_PREVISTA" value="${Utils.dateValue(r.DATA_PREVISTA)}">
        </div>
        <div class="form-group">
          <label>Data Conclusão</label>
          <input type="date" id="me-DATA_CONCLUSAO" value="${Utils.dateValue(r.DATA_CONCLUSAO)}">
        </div>
        <div class="form-group">
          <label>Status</label>
          <select id="me-STATUS">
            <option value="">Em Andamento</option>
            <option value="OK"${r.STATUS === 'OK' ? ' selected' : ''}>OK (Concluído)</option>
          </select>
        </div>
        <div class="form-group form-full">
          <label>Ação de Controle *</label>
          <textarea id="me-ACAO_CONTROLE" rows="4">${Utils.esc(r.ACAO_CONTROLE)}</textarea>
        </div>
        <div class="form-group form-full">
          <label>Observações</label>
          <textarea id="me-OBSERVACOES" rows="3">${Utils.esc(r.OBSERVACOES)}</textarea>
        </div>
        <div class="form-group">
          <label>Eficácia</label>
          <input type="text" id="me-EFICACIA" value="${Utils.esc(r.EFICACIA)}">
        </div>
        <div class="form-actions form-full">
          <button type="button" class="btn-secondary" onclick="Modal.close()">Cancelar</button>
          <button type="button" class="btn-primary" onclick="Modal.savePA()">Salvar</button>
        </div>
      </div>`);
  },

  async savePA() {
    const get = id => (document.getElementById('me-' + id) || {}).value || '';
    const data = {
      ...State.editTarget.orig,
      SETOR:             get('SETOR'),
      POSTO_TRABALHO:    get('POSTO_TRABALHO'),
      CRITICIDADE:       get('CRITICIDADE'),
      CLASSIFICACAO:     get('CLASSIFICACAO'),
      GERENTE:           get('GERENTE'),
      RESPONSAVEL:       get('RESPONSAVEL'),
      ESTIMATIVA_VALOR:  get('ESTIMATIVA_VALOR'),
      DATA_PREVISTA:     get('DATA_PREVISTA'),
      DATA_CONCLUSAO:    get('DATA_CONCLUSAO'),
      STATUS:            get('STATUS'),
      ACAO_CONTROLE:     get('ACAO_CONTROLE'),
      OBSERVACOES:       get('OBSERVACOES'),
      EFICACIA:          get('EFICACIA')
    };
    if (!data.SETOR) { Utils.toast('Informe o setor.', 'error'); return; }
    try {
      await API.update(CONFIG.SHEETS.PA, State.editTarget.rowNum, data);
      Utils.toast('Ação atualizada com sucesso!', 'success');
      this.close();
      await App.loadPA();
    } catch (e) {
      Utils.toast('Erro ao salvar: ' + e.message, 'error');
    }
  }
};

// ================================================================
//  FORMS (novos registros)
// ================================================================
const Forms = {
  async submitAET(e) {
    e.preventDefault();
    const btn = document.getElementById('btn-submit-aet');
    const data = Object.fromEntries(new FormData(e.target).entries());
    btn.disabled = true; btn.textContent = 'Salvando…';
    try {
      await API.create(CONFIG.SHEETS.AET, data);
      Utils.toast('Posto criado com sucesso!', 'success');
      e.target.reset();
      await App.loadAET();
    } catch (err) {
      Utils.toast('Erro ao salvar: ' + err.message, 'error');
    } finally {
      btn.disabled = false; btn.textContent = 'Salvar Posto';
    }
  },

  async submitPA(e) {
    e.preventDefault();
    const btn = document.getElementById('btn-submit-pa');
    const data = Object.fromEntries(new FormData(e.target).entries());
    btn.disabled = true; btn.textContent = 'Salvando…';
    try {
      await API.create(CONFIG.SHEETS.PA, data);
      Utils.toast('Ação criada com sucesso!', 'success');
      e.target.reset();
      await App.loadPA();
    } catch (err) {
      Utils.toast('Erro ao salvar: ' + err.message, 'error');
    } finally {
      btn.disabled = false; btn.textContent = 'Salvar Ação';
    }
  }
};

// ================================================================
//  NAV SETUP
// ================================================================
function setupNav() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(s => s.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
    });
  });

  document.querySelectorAll('.lanc-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.lanc-tab').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.form-card').forEach(f => f.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.form).classList.add('active');
    });
  });

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') Modal.close();
  });
}

// ================================================================
//  APP
// ================================================================
const App = {
  async loadAET() {
    try {
      const data = await API.read(CONFIG.SHEETS.AET);
      AET.load(data);
    } catch (e) {
      document.getElementById('aet-tbody').innerHTML =
        `<tr><td colspan="8" class="table-loading">Erro ao carregar: ${Utils.esc(e.message)}</td></tr>`;
    }
  },

  async loadPA() {
    try {
      const data = await API.read(CONFIG.SHEETS.PA);
      PA.load(data);
    } catch (e) {
      document.getElementById('pa-tbody').innerHTML =
        `<tr><td colspan="8" class="table-loading">Erro ao carregar: ${Utils.esc(e.message)}</td></tr>`;
    }
  },

  async refresh() {
    const btn = document.getElementById('btnRefresh');
    btn.classList.add('loading'); btn.disabled = true;
    await Promise.all([this.loadAET(), this.loadPA()]);
    btn.classList.remove('loading'); btn.disabled = false;
    document.getElementById('lastUpdate').textContent =
      'Atualizado às ' + new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
  },

  init() {
    setupNav();
    if (!API.isConfigured()) {
      document.getElementById('setup-banner').style.display = 'flex';
      document.getElementById('aet-tbody').innerHTML =
        '<tr><td colspan="8" class="table-loading">Configure a URL do Apps Script em app.js para carregar os dados.</td></tr>';
      document.getElementById('pa-tbody').innerHTML =
        '<tr><td colspan="8" class="table-loading">Configure a URL do Apps Script em app.js para carregar os dados.</td></tr>';
      return;
    }
    document.getElementById('setup-banner').style.display = 'none';
    this.refresh();
  }
};

document.addEventListener('DOMContentLoaded', () => App.init());
