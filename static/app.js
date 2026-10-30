// ============================================
// DEBTMAIL PRO - MAIL MERGE APP LOGIC
// ============================================

const App = {
  customers: [],
  selectedCustomer: null,
  currentPreviewHtml: '',
  checkedCustomers: new Set(),
  sentCustomers: new Set(),
  settings: {},
  globalAttachments: [], // [{name, path}]
  customerAttachments: {}, // { 'customer_name': [{name, path}] }
  sendLog: {}, // { 'customer_name': { count, history } }
  outlookAccounts: [], // [{email, name}]

  async init() {
    await this.loadSettings();
    await this.loadCustomers();
    await this.loadSendLog();
    this.checkOutlook();
    this.bindEvents();
  },

  // ---- API ----
  async api(url, opts = {}) {
    try {
      const res = await fetch(url, opts);
      return await res.json();
    } catch (e) {
      this.toast('Server error: ' + e.message, 'error');
      return null;
    }
  },

  async loadSettings() {
    const data = await this.api('/api/settings');
    if (data) {
      this.settings = data;
      this.fillSettingsForm();
      const input = document.getElementById('kh_attachments_dir');
      if (input) {
        input.value = data.kh_attachments_dir || '';
      }
    }
  },

  async saveSettings() {
    const s = this.readSettingsForm();
    const data = await this.api('/api/settings', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(s)
    });
    if (data?.success) {
      this.settings = data.settings;
      this.toast('Cài đặt đã lưu', 'success');
      this.closeModal('settingsModal');
      const input = document.getElementById('kh_attachments_dir');
      if (input) {
        input.value = this.settings.kh_attachments_dir || '';
      }
      // Refresh From dropdowns
      const fromSelect = document.getElementById('fieldFrom');
      if (fromSelect) {
        fromSelect.innerHTML = this._buildAccountOptions();
        if (this.settings.sender_email) {
          fromSelect.value = this.settings.sender_email;
        }
      }
      // Re-generate preview if a customer is selected
      if (this.selectedCustomer) this.generatePreview();
    }
  },

  // Build <option> HTML for all accounts + aliases
  _buildAccountOptions() {
    let html = '';
    for (const a of this.outlookAccounts) {
      const aliases = a.aliases || [];
      if (aliases.length > 0) {
        html += `<optgroup label="${a.name}">`;
        html += `<option value="${a.email}">📧 ${a.email}</option>`;
        for (const alias of aliases) {
          html += `<option value="${alias}">↳ ${alias} [Alias]</option>`;
        }
        html += `</optgroup>`;
      } else {
        html += `<option value="${a.email}">${a.name} &lt;${a.email}&gt;</option>`;
      }
    }

    // Manual aliases
    const manualStr = this.settings.manual_aliases || '';
    if (manualStr.trim()) {
      const manualList = manualStr.split('\n')
        .map(s => s.trim())
        .filter(s => s.length > 0 && s.includes('@'));
      if (manualList.length > 0) {
        html += `<optgroup label="Alias nhập thủ công">`;
        for (const alias of manualList) {
          html += `<option value="${alias}">↳ ${alias} [Alias thủ công]</option>`;
        }
        html += `</optgroup>`;
      }
    }
    return html;
  },

  async checkOutlook() {
    const data = await this.api('/api/outlook-check');
    const dot = document.getElementById('outlookStatus');
    const txt = document.getElementById('outlookStatusText');
    const fromSelect = document.getElementById('fieldFrom');
    const setSenderEmail = document.getElementById('set_sender_email');

    if (data?.available) {
      dot.className = 'status-dot online';
      txt.textContent = 'Outlook: Connected';
      this.outlookAccounts = data.accounts || [];

      const optsHtml = this._buildAccountOptions();
      if (fromSelect) fromSelect.innerHTML = optsHtml;
      if (setSenderEmail) {
        setSenderEmail.innerHTML = optsHtml;
        if (this.settings.sender_email) {
          setSenderEmail.value = this.settings.sender_email;
        }
      }

      // Auto-select saved account
      const saved = this.settings.sender_email;
      if (saved && fromSelect) {
        fromSelect.value = saved;
      }
      // If nothing selected, use first account
      if (fromSelect && !fromSelect.value && this.outlookAccounts.length) {
        fromSelect.value = this.outlookAccounts[0].email;
      }
    } else {
      dot.className = 'status-dot offline';
      const errMsg = data?.error ? ` (${data.error})` : '';
      txt.textContent = 'Outlook: Offline' + errMsg;
      if (fromSelect) fromSelect.innerHTML = '<option value="">Outlook chưa kết nối</option>';
      if (setSenderEmail) setSenderEmail.innerHTML = '<option value="">Outlook chưa kết nối</option>';
    }
  },

  async loadCustomers() {
    const data = await this.api('/api/customers');
    if (data?.customers) {
      this.customers = data.customers;
      this.renderCustomerList();
    }
  },

  // ---- SELECT CUSTOMER ----
  async selectCustomer(name) {
    const data = await this.api('/api/customer/' + encodeURIComponent(name));
    if (!data?.success) return;

    this.selectedCustomer = data;
    document.querySelectorAll('.customer-item').forEach(el => {
      el.classList.toggle('active', el.dataset.name === name);
    });

    const cust = data.customer;
    const s = this.settings;

    document.getElementById('fieldFrom').value = s.sender_email || '';
    document.getElementById('fieldTo').value = cust.email_to || '';
    document.getElementById('fieldCc').value = cust.email_cc || '';
    document.getElementById('fieldBcc').value = cust.email_bcc || '';

    // Generate subject with debt_date from Excel
    const subjData = await this.api('/api/generate-subject', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        short_name: cust.short_name || '',
        debt_date: cust.debt_date || '',
        format: s.subject_format || ''
      })
    });
    document.getElementById('fieldSubject').value = subjData?.subject || '';

    this.renderAttachments();
    this.renderSendLog();
    await this.generatePreview();
  },

  async generatePreview() {
    if (!this.selectedCustomer) return;
    const cust = this.selectedCustomer.customer;
    const s = this.settings;

    const data = await this.api('/api/preview', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        template: s.email_template,
        signature: s.signature,
        customer: cust,
        debt_items: this.selectedCustomer.debt_items
      })
    });

    if (data?.html) {
      this.currentPreviewHtml = data.html;
      const frame = document.getElementById('emailPreviewContent');
      frame.innerHTML = data.html;
      document.getElementById('emailEmpty').style.display = 'none';
      frame.style.display = 'block';
      document.getElementById('editToolbar').classList.add('visible');
    }
  },

  getAttachmentsFor(customerName, matched = []) {
    const custAtts = this.customerAttachments[customerName] || [];
    return [...this.globalAttachments, ...custAtts, ...matched];
  },

  // ---- SEND ----
  async sendCurrentEmail() {
    if (!this.selectedCustomer) {
      this.toast('Chưa chọn khách hàng', 'error');
      return;
    }
    const to = document.getElementById('fieldTo').value;
    const cc = document.getElementById('fieldCc').value;
    const bcc = document.getElementById('fieldBcc').value;
    const subject = document.getElementById('fieldSubject').value;
    const sendMode = document.querySelector('input[name="sendMode"]:checked')?.value || 'display';
    const editedHtml = document.getElementById('emailPreviewContent').innerHTML;
    const currentName = this.selectedCustomer.customer.customer_name;
    const matched = this.selectedCustomer.matched_attachments || [];

    const data = await this.api('/api/send', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        to, cc, bcc, subject,
        html_body: editedHtml,
        customer_name: currentName,
        attachments: this.getAttachmentsFor(currentName, matched).map(a => a.path),
        send_mode: sendMode,
        from_account: document.getElementById('fieldFrom').value
      })
    });

    if (data?.success) {
      this.toast(data.message, 'success');
      this.sentCustomers.add(this.selectedCustomer.customer.customer_name);
      // Update log
      await this.loadSendLog();
      this.renderSendLog();
      this.renderCustomerList();
    } else {
      this.toast(data?.message || 'Send error', 'error');
    }
  },

  async sendBatch() {
    if (this.checkedCustomers.size === 0) {
      this.toast('Chưa chọn khách hàng nào', 'error');
      return;
    }
    const sendMode = document.querySelector('input[name="sendMode"]:checked')?.value || 'display';
    const s = this.settings;
    const emails = [];

    for (const name of this.checkedCustomers) {
      const detail = await this.api('/api/customer/' + encodeURIComponent(name));
      if (!detail?.success) continue;

      const cust = detail.customer;
      const subjData = await this.api('/api/generate-subject', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          short_name: cust.short_name || '',
          debt_date: cust.debt_date || '',
          format: s.subject_format || ''
        })
      });

      const preview = await this.api('/api/preview', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          template: s.email_template,
          signature: s.signature,
          customer: cust,
          debt_items: detail.debt_items
        })
      });

      const matched = detail.matched_attachments || [];
      emails.push({
        customer_name: name,
        to: cust.email_to || '',
        cc: cust.email_cc || '',
        bcc: cust.email_bcc || '',
        subject: subjData?.subject || '',
        html_body: preview?.html || '',
        attachments: this.getAttachmentsFor(name, matched).map(a => a.path),
        from_account: s.sender_email
      });
    }

    const result = await this.api('/api/send-batch', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ emails, send_mode: sendMode })
    });

    if (result?.success) {
      this.toast(result.message, 'success');
      await this.loadSendLog();
      result.results?.forEach(r => {
        if (r.success) this.sentCustomers.add(r.customer);
      });
      this.renderCustomerList();
    }
  },

  async uploadExcel(file) {
    const formData = new FormData();
    formData.append('file', file);
    const data = await this.api('/api/upload', { method: 'POST', body: formData });
    if (data?.success) {
      this.toast(data.message, 'success');
      await this.loadCustomers();
    } else {
      this.toast(data?.message || 'Upload error', 'error');
    }
  },

  async uploadErp(file) {
    const choice = confirm('Bạn muốn:\n- Bấm OK: Nạp thẳng dữ liệu vào Tool\n- Bấm Cancel: Chỉ xuất file Excel Template để tải về');
    const mode = choice ? 'import' : 'download';

    this.toast(`Đang xử lý file ERP (${mode})...`, 'info');
    const formData = new FormData();
    formData.append('file', file);
    formData.append('output_mode', mode);

    if (mode === 'download') {
      try {
        const res = await fetch('/api/upload-erp', { method: 'POST', body: formData });
        if (!res.ok) throw new Error('Server error');
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = 'Template_Tu_ERP.xlsx';
        document.body.appendChild(a); a.click();
        document.body.removeChild(a); URL.revokeObjectURL(url);
        this.toast('Đã tải file chuyển đổi xong!', 'success');
      } catch (e) { this.toast('Lỗi tải file: ' + e.message, 'error'); }
    } else {
      const data = await this.api('/api/upload-erp', { method: 'POST', body: formData });
      if (data?.success) {
        this.toast(data.message, 'success');
        await this.loadCustomers();
      } else {
        this.toast(data?.message || 'Convert error', 'error');
      }
    }
    // Reset file input
    const el = document.getElementById('fileInputErp') || document.getElementById('fileInputB8');
    if (el) el.value = '';
  },
  uploadB8(file) { return this.uploadErp(file); },

  async clearData() {
    if (!confirm('Bạn có chắc muốn xóa toàn bộ danh sách hiện tại?')) return;
    const data = await this.api('/api/clear-data', { method: 'POST' });
    if (data?.success) {
      this.toast(data.message, 'success');
      this.customers = [];
      this.selectedCustomer = null;
      this.checkedCustomers.clear();
      this.sentCustomers.clear();
      this.renderCustomerList();
      
      // Reset preview
      document.getElementById('emailPreviewContent').style.display = 'none';
      document.getElementById('emailEmpty').style.display = 'flex';
      document.getElementById('editToolbar').classList.remove('visible');
    }
  },

  // ---- RENDER ----
  renderCustomerList(filter = '') {
    const list = document.getElementById('customerList');
    const filtered = this.customers.filter(c =>
      !filter || c.customer_name.toLowerCase().includes(filter.toLowerCase()) ||
      (c.short_name || '').toLowerCase().includes(filter.toLowerCase())
    );

    if (filtered.length === 0) {
      list.innerHTML = `<div class="email-empty" style="padding:40px 0;">
        <span class="icon">📋</span>
        <p>${this.customers.length === 0 ? 'Upload file Excel để bắt đầu' : 'Không tìm thấy'}</p>
      </div>`;
      return;
    }

    list.innerHTML = filtered.map(c => {
      const name = c.customer_name;
      const checked = this.checkedCustomers.has(name) ? 'checked' : '';
      const active = this.selectedCustomer?.customer?.customer_name === name ? 'active' : '';
      const sent = this.sentCustomers.has(name) ? 'sent' : '';
      const badge = this.sentCustomers.has(name)
        ? '<span class="badge badge-sent">Đã gửi</span>'
        : '<span class="badge badge-pending">Chưa gửi</span>';
      const debt = c.total_debt ? Number(c.total_debt).toLocaleString('vi-VN') : '0';
      const logEntry = this.sendLog[name];
      const remindBadge = logEntry?.count > 0 
        ? `<span class="badge" style="background:#e8eaf6;color:#3949ab">🔔${logEntry.count}</span>` 
        : '';

      return `<div class="customer-item ${active} ${sent}" data-name="${name}" onclick="App.selectCustomer('${name.replace(/'/g, "\\'")}')">
        <div class="customer-check">
          <input type="checkbox" ${checked} onclick="event.stopPropagation(); App.toggleCheck('${name.replace(/'/g, "\\'")}', this.checked)">
        </div>
        <div class="customer-info">
          <div class="customer-name" title="${name}">${name}</div>
          <div class="customer-meta">
            <span class="customer-debt">₫${debt}</span>
            <span>${c.debt_count || 0} khoản</span>
            ${badge}
            ${remindBadge}
          </div>
        </div>
      </div>`;
    }).join('');

    document.getElementById('customerCount').textContent = `${filtered.length} KH`;
  },

  toggleCheck(name, checked) {
    if (checked) this.checkedCustomers.add(name);
    else this.checkedCustomers.delete(name);
    this.updateBatchCount();
  },

  selectAll() {
    this.customers.forEach(c => this.checkedCustomers.add(c.customer_name));
    this.renderCustomerList(document.getElementById('searchInput').value);
    this.updateBatchCount();
  },

  deselectAll() {
    this.checkedCustomers.clear();
    this.renderCustomerList(document.getElementById('searchInput').value);
    this.updateBatchCount();
  },

  updateBatchCount() {
    document.getElementById('btnSendBatch').textContent = `📨 Gửi tất cả đã chọn (${this.checkedCustomers.size})`;
  },

  // ---- SETTINGS ----
  fillSettingsForm() {
    const s = this.settings;
    // Populate sender dropdown from detected accounts + aliases
    const sel = document.getElementById('set_sender_email');
    if (this.outlookAccounts?.length) {
      sel.innerHTML = this._buildAccountOptions();
    }
    sel.value = s.sender_email || '';
    document.getElementById('set_subject_format').value = s.subject_format || '';
    document.getElementById('set_email_template').value = s.email_template || '';
    document.getElementById('set_signature').value = s.signature || '';
    document.getElementById('set_kh_attachments_dir').value = s.kh_attachments_dir || '';
    document.getElementById('set_manual_aliases').value = s.manual_aliases || '';
  },

  readSettingsForm() {
    return {
      sender_email: document.getElementById('set_sender_email').value,
      subject_format: document.getElementById('set_subject_format').value,
      email_template: document.getElementById('set_email_template').value,
      signature: document.getElementById('set_signature').value,
      kh_attachments_dir: document.getElementById('set_kh_attachments_dir').value,
      manual_aliases: document.getElementById('set_manual_aliases').value,
    };
  },

  // ---- ATTACHMENTS (File Picker) ----
  addAttachment(isGlobal) {
    if (!isGlobal && !this.selectedCustomer) {
      this.toast('Vui lòng chọn 1 khách hàng trước!', 'error');
      return;
    }
    if (window.showOpenFilePicker) {
      this._addAttachmentModern(isGlobal);
    } else {
      const path = prompt('Nhập đường dẫn file đính kèm:');
      if (path) {
        this._pushAttachment(isGlobal, { name: path.split('\\').pop().split('/').pop(), path: path });
        this.renderAttachments();
      }
    }
  },

  async _addAttachmentModern(isGlobal) {
    try {
      const handles = await window.showOpenFilePicker({ multiple: true });
      for (const handle of handles) {
        const file = await handle.getFile();
        this._pushAttachment(isGlobal, { name: file.name, path: file.name, handle: handle });
      }
      this.renderAttachments();
      this.toast(`Đã thêm ${handles.length} file đính kèm`, 'success');
    } catch (e) {
      if (e.name !== 'AbortError') {
        const path = prompt('Nhập đường dẫn đầy đủ file đính kèm\n(VD: D:\\Documents\\report.pdf):');
        if (path) {
          this._pushAttachment(isGlobal, { name: path.split('\\').pop().split('/').pop(), path: path });
          this.renderAttachments();
        }
      }
    }
  },

  _pushAttachment(isGlobal, obj) {
    if (isGlobal) {
      this.globalAttachments.push(obj);
    } else {
      const name = this.selectedCustomer.customer.customer_name;
      if (!this.customerAttachments[name]) this.customerAttachments[name] = [];
      this.customerAttachments[name].push(obj);
    }
  },

  removeAttachment(isGlobal, idx) {
    if (isGlobal) {
      this.globalAttachments.splice(idx, 1);
    } else {
      const name = this.selectedCustomer.customer.customer_name;
      if (this.customerAttachments[name]) {
        this.customerAttachments[name].splice(idx, 1);
      }
    }
    this.renderAttachments();
  },

  renderAttachments() {
    const container = document.getElementById('attachmentList');
    const globals = this.globalAttachments;
    const custs = this.selectedCustomer?.customer?.customer_name ? 
      (this.customerAttachments[this.selectedCustomer.customer.customer_name] || []) : [];
    const matched = this.selectedCustomer?.matched_attachments || [];

    if (globals.length === 0 && custs.length === 0 && matched.length === 0) {
      container.innerHTML = '<span style="font-size:12px;color:var(--text-muted)">Chưa có file đính kèm</span>';
      return;
    }

    let html = '';
    if (globals.length > 0) {
      html += '<div style="font-size:10px;font-weight:bold;color:var(--text-muted);margin:8px 0 4px;">FILE CHUNG MỌI KHÁCH HÀNG:</div>';
      html += globals.map((a, i) =>
        `<div class="attachment-item">
          <span style="font-size:10px;padding:2px 4px;background:#e8f0fe;color:var(--accent);border-radius:3px;">Chung</span>
          <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin-left:6px;" title="${a.path}">${a.name}</span>
          <button class="remove-att" onclick="App.removeAttachment(true, ${i})">✕</button>
        </div>`
      ).join('');
    }
    
    if (custs.length > 0) {
      html += '<div style="font-size:10px;font-weight:bold;color:var(--text-muted);margin:8px 0 4px;">FILE RIÊNG CHO KH NÀY:</div>';
      html += custs.map((a, i) =>
        `<div class="attachment-item" style="border-left:3px solid #ff9800;">
          <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin-left:6px;" title="${a.path}">${a.name}</span>
          <button class="remove-att" onclick="App.removeAttachment(false, ${i})">✕</button>
        </div>`
      ).join('');
    }

    if (matched.length > 0) {
      html += '<div style="font-size:10px;font-weight:bold;color:var(--success);margin:8px 0 4px;">FILE TỰ ĐỘNG KHỚP (MÃ KH):</div>';
      html += matched.map((a, i) =>
        `<div class="attachment-item matched-att-item">
          <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin-left:6px;" title="${a.path}">${a.name}</span>
        </div>`
      ).join('');
    }
    container.innerHTML = html;
  },

  // ---- MERGE TAG COPY ----
  copyTag(tag) {
    navigator.clipboard.writeText(tag).then(() => {
      this.toast(`Đã copy: ${tag}`, 'info');
    });
  },

  // ---- DOWNLOAD TEMPLATE ----
  async downloadTemplate() {
    try {
      this.toast('Đang tạo template...', 'info');
      const response = await fetch('/api/sample-excel');
      if (!response.ok) throw new Error('Server error');
      const blob = await response.blob();

      if (window.showSaveFilePicker) {
        try {
          const handle = await window.showSaveFilePicker({
            suggestedName: 'template_cong_no.xlsx',
            types: [{ description: 'Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
          });
          const w = await handle.createWritable();
          await w.write(blob);
          await w.close();
          this.toast('Template đã lưu!', 'success');
          return;
        } catch (e) { if (e.name === 'AbortError') return; }
      }

      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'template_cong_no.xlsx';
      document.body.appendChild(a); a.click();
      document.body.removeChild(a); URL.revokeObjectURL(url);
      this.toast('Template đã tải!', 'success');
    } catch (e) {
      this.toast('Lỗi: ' + e.message, 'error');
    }
  },

  // ---- SEND LOG ----
  async loadSendLog() {
    const data = await this.api('/api/send-log');
    if (data) this.sendLog = data;
  },

  renderSendLog() {
    const el = document.getElementById('sendLogSummary');
    if (!el) return;
    if (!this.selectedCustomer) {
      el.innerHTML = '<span style="color:var(--text-muted)">Chọn KH để xem lịch sử</span>';
      return;
    }
    const name = this.selectedCustomer.customer.customer_name;
    const log = this.sendLog[name];
    if (!log || log.count === 0) {
      el.innerHTML = '<span style="color:var(--text-muted)">Chưa có lịch sử nhắc nợ</span>';
      return;
    }
    let html = `<div style="margin-bottom:8px"><b style="color:var(--accent);font-size:14px">${log.count}</b> lần đã nhắc</div>`;
    const recent = (log.history || []).slice(-5).reverse();
    html += recent.map(h => 
      `<div style="font-size:10px;padding:4px 0;border-bottom:1px solid #eee;color:var(--text-secondary)">
        📅 ${h.date}<br>
        <span style="color:var(--text-muted)">${h.to}</span>
      </div>`
    ).join('');
    el.innerHTML = html;
  },

  // ---- EDIT GUIDE ----
  toggleEditGuide() {
    const guide = document.getElementById('editGuide');
    if (guide) {
      guide.style.display = guide.style.display === 'none' ? 'block' : 'none';
    }
  },

  switchTab(paneId, btnEl) {
    // Hide all panes inside the current container
    const container = btnEl.closest('.tab-container');
    container.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
    // Show selected pane
    container.querySelector('#' + paneId).classList.add('active');
    // Deactivate all buttons in this header
    container.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    // Activate clicked button
    btnEl.classList.add('active');
  },

  // ---- MODALS ----
  openModal(id) { 
    document.getElementById(id).classList.add('active');
    if (id === 'aboutModal') this.loadLicenseInfo();
  },
  closeModal(id) { document.getElementById(id).classList.remove('active'); },

  async loadLicenseInfo() {
    const el = document.getElementById('licenseInfo');
    if (!el) return;
    const data = await this.api('/api/license');
    if (!data || !data.activated) {
      el.innerHTML = '⏳ Chưa kích hoạt';
      el.style.background = '#fff3e0'; el.style.color = '#e65100';
      return;
    }
    if (data.expired) {
      el.innerHTML = `❌ Đã hết hạn (${data.expiry}). Vui lòng liên hệ quản trị viên để gia hạn.`;
      el.style.background = '#ffebee'; el.style.color = '#c62828';
    } else if (data.remaining <= 30) {
      el.innerHTML = `⚠️ Còn ${data.remaining} ngày | Hết hạn: ${data.expiry}`;
      el.style.background = '#fff3e0'; el.style.color = '#e65100';
    } else {
      el.innerHTML = `✅ Kích hoạt: ${data.activated.split(' ')[0]} | Hết hạn: ${data.expiry} | Còn ${data.remaining} ngày`;
      el.style.background = '#e8f5e9'; el.style.color = '#2e7d32';
    }
  },

  // ---- TOASTS ----
  toast(msg, type = 'info') {
    const container = document.getElementById('toastContainer');
    const t = document.createElement('div');
    t.className = 'toast toast-' + type;
    t.innerHTML = `<span>${type === 'success' ? '✓' : type === 'error' ? '✗' : 'ℹ'}</span> ${msg}`;
    container.appendChild(t);
    setTimeout(() => { t.style.opacity = '0'; setTimeout(() => t.remove(), 300); }, 4000);
  },

  // ---- EVENTS ----
  bindEvents() {
    document.getElementById('searchInput').addEventListener('input', e => {
      this.renderCustomerList(e.target.value);
    });
    document.getElementById('fileInput').addEventListener('change', e => {
      if (e.target.files[0]) this.uploadExcel(e.target.files[0]);
    });
    const erpInput = document.getElementById('fileInputErp') || document.getElementById('fileInputB8');
    if (erpInput) {
      erpInput.addEventListener('change', e => {
        if (e.target.files[0]) this.uploadErp(e.target.files[0]);
      });
    }
    document.getElementById('kh_attachments_dir').addEventListener('change', async e => {
      this.settings.kh_attachments_dir = e.target.value;
      await this.api('/api/settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(this.settings)
      });
      const setInput = document.getElementById('set_kh_attachments_dir');
      if (setInput) setInput.value = e.target.value;
      if (this.selectedCustomer) {
        await this.selectCustomer(this.selectedCustomer.customer.customer_name);
      }
    });
  }
};

document.addEventListener('DOMContentLoaded', () => App.init());
