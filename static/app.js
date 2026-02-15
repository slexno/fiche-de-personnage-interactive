let state = null;
let modalInventoryId = null;
let modalShopRef = null;

const money = (v) => Number(v || 0).toFixed(2);
const clean = (v) => (v === undefined || v === null ? '' : String(v));
const isFilled = (v) => clean(v).trim() !== '' && clean(v).trim() !== '0';

const apiAction = async (payload) => {
  const res = await fetch('/api/action', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });
  const json = await res.json();
  if (json.ok === false && payload.action === 'buy') {
    const miss = money(json.missing_credits || 0);
    showAlertModal(`Fonds insuffisants. Il vous manque ${miss} crédits.`);
  }
  state = json.state;
  render();
  return json;
};

const init = async () => {
  const res = await fetch('/api/state');
  state = await res.json();
  bindTabs();
  render();
};

const bindTabs = () => {
  document.querySelectorAll('.tabs button').forEach(btn => {
    btn.onclick = () => {
      document.querySelectorAll('.tabs button').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.tab).classList.add('active');
    };
  });
};

const render = () => {
  renderStats();
  renderInventory();
  renderShop();
};

const renderStats = () => {
  const s = state.stats;
  document.getElementById('stats').innerHTML = `
    <div class="grid">
      <div class="panel">
        <h2>Statistiques</h2>
        <div class="table-wrap"><table>
          <tr><th>Nom</th><th>Score</th><th>Bonus</th></tr>
          ${s.stats.map(st => `<tr>
            <td>${st.name}</td>
            <td><input type="number" min="1" max="20" value="${st.score}" onchange="updateStat('${st.name}', this.value)"></td>
            <td>${st.bonus >= 0 ? '+' : ''}${st.bonus}</td>
          </tr>`).join('')}
        </table></div>
        <p><strong>Armor class :</strong> ${s.armor_class}</p>
      </div>
      <div class="panel">
        <h2>Compétences</h2>
        <div class="table-wrap"><table>
          <tr><th>Compétence</th><th>Modif</th><th>Bonus</th><th>Spécialisation</th></tr>
          ${s.skills.map(sk => `<tr>
            <td>${sk.name}</td><td>${sk.mod}</td><td>${sk.bonus >= 0 ? '+' : ''}${sk.bonus}</td>
            <td><input type="checkbox" ${sk.specialized ? 'checked' : ''} onchange="toggleSkill('${sk.name}', this.checked)"></td>
          </tr>`).join('')}
        </table></div>
      </div>
    </div>`;
};

const itemRow = (it, from) => {
  const to = from === 'sac à dos' ? 'coffre' : 'sac à dos';
  return `<tr class="clickable" onclick="openInventoryModal('${it.id}')">
    <td>${it.Objet || ''}</td>
    <td><input type="number" step="0.1" value="${it['Prix unitaire (en crédit)'] || 0}" onchange="quickUpdate('${it.id}','Prix unitaire (en crédit)',this.value)" onclick="event.stopPropagation()"></td>
    <td><input type="number" step="0.1" value="${it['poid unitaire (kg)'] || 0}" onchange="quickUpdate('${it.id}','poid unitaire (kg)',this.value)" onclick="event.stopPropagation()"></td>
    <td><input type="number" min="0" value="${it['Quantité'] || 1}" onchange="quickUpdate('${it.id}','Quantité',this.value)" onclick="event.stopPropagation()"></td>
    <td>${it.type || 'item'}</td>
    <td>
      <input id="move-${it.id}" type="number" min="1" max="${it['Quantité'] || 1}" value="1" style="width:70px" onclick="event.stopPropagation()">
      <button onclick="event.stopPropagation(); transferQty('${from}','${to}','${it.id}')">Transférer</button>
    </td>
  </tr>`;
};

const renderInventory = () => {
  const inv = state.inventory;
  document.getElementById('inventory').innerHTML = `
  <div class="grid">
    <div class="panel">
      <h2>Sac à dos</h2>
      <div class="row"><strong>Crédits:</strong> ${money(inv.credits)} | <strong>Poids:</strong> ${money(inv.bag_weight)}kg ${inv.overweight ? '<span class="warn">⚠ surcharge -1 dex</span>' : ''}</div>
      <div class="row">
        <button onclick="sortBag('alpha')">A-Z</button>
        <button onclick="sortBag('prix')">Prix</button>
        <button onclick="sortBag('poids')">Poids</button>
      </div>
      <div class="table-wrap"><table>
        <tr><th>Objet</th><th>Valeur unitaire</th><th>Poids unitaire</th><th>Quantité</th><th>Type</th><th>Action</th></tr>
        ${inv.bag.map(i => itemRow(i, 'sac à dos')).join('')}
      </table></div>

      <h3>Ajouter manuellement</h3>
      <div class="row">
        <input id="n" placeholder="Nom">
        <input id="pu" type="number" step="0.1" placeholder="Prix unitaire">
        <input id="wu" type="number" step="0.1" placeholder="Poids unitaire">
        <input id="q" type="number" value="1" placeholder="Quantité">
        <input id="d" placeholder="Description">
        <input id="range" placeholder="Range">
        <input id="hit" placeholder="Hit">
        <input id="damage" placeholder="Damage">
        <input id="ac" placeholder="Armor class">
        <input id="effet" placeholder="Effet">
        <button onclick="addItem()">Ajouter</button>
      </div>
    </div>

    <div class="panel">
      <h2>Coffre</h2>
      <div class="table-wrap"><table>
        <tr><th>Objet</th><th>Valeur unitaire</th><th>Poids unitaire</th><th>Quantité</th><th>Type</th><th>Action</th></tr>
        ${inv.chest.map(i => itemRow(i, 'coffre')).join('')}
      </table></div>

      <h3>Tableau des armes</h3>
      <div class="table-wrap"><table>
        <tr><th>Nom</th><th>Range</th><th>Hit</th><th>Damage</th><th>Équipé</th></tr>
        ${inv.weapons.map(w => `<tr><td>${w.Objet || ''}</td><td>${w['Range (ft)'] || '-'}</td><td>${w.Hit || '-'}</td><td>${w.Damage || '-'}</td><td>${w.equiped === '1' ? 'Oui' : 'Non'}</td></tr>`).join('')}
      </table></div>

      <h3>Tableau des équipements</h3>
      <div class="table-wrap"><table>
        <tr><th>Nom</th><th>Armor class</th><th>Effet</th><th>Équipé</th></tr>
        ${inv.equipments.map(e => `<tr><td>${e.Objet || ''}</td><td>${e['bonus Armor class'] || '0'}</td><td>${e['effet(optionel)'] || '-'}</td><td>${e.equiped === '1' ? 'Oui' : 'Non'}</td></tr>`).join('')}
      </table></div>
      <p class="small">Cliquez un objet pour ouvrir le popup (édition, assignation, vente).</p>
    </div>
  </div>`;
};

const renderShop = () => {
  const sections = Object.entries(state.shop).map(([sheet, items]) => `
    <div class="panel">
      <div class="row"><h3>${sheet}</h3><span class="credit-badge">💳 Crédits: ${money(state.inventory.credits)}</span></div>
      <div class="table-wrap"><table>
        <tr><th>Objet</th><th>Prix</th><th>Poids</th><th>Description</th><th>Achat</th></tr>
        ${items.map(i => {
          const name = i["nom de l'objet"] || '';
          return `<tr class="clickable" onclick="openShopModal('${sheet}','${encodeURIComponent(name)}')">
            <td>${name}</td><td>${money(i['prix unitaire (crédit)'])}</td><td>${money(i['poid unitaire(kg)'])}</td><td>${i.description || ''}</td>
            <td><input id="buy-${sheet}-${name.replace(/\s+/g,'_')}" type="number" value="1" onclick="event.stopPropagation()" style="width:70px"><button onclick="event.stopPropagation(); buy('${sheet}',\`${name}\`)">Acheter</button></td>
          </tr>`;
        }).join('')}
      </table></div>
    </div>`).join('');
  document.getElementById('shop').innerHTML = sections;
};

const getInventoryItem = (id) => [...state.inventory.bag, ...state.inventory.chest].find(x => x.id === id);

window.updateStat = (name, score) => apiAction({ action: 'update_stat', name, score });
window.toggleSkill = (name, specialized) => apiAction({ action: 'toggle_skill', name, specialized });
window.sortBag = (key) => apiAction({ action: 'sort', key, source: 'sac à dos' });
window.transferQty = (from, to, id) => apiAction({ action: 'transfer_item', from, to, id, qty: document.getElementById(`move-${id}`).value });
window.quickUpdate = (id, key, value) => apiAction({ action: 'update_item', id, [key]: value });
window.buy = (sheet, name, qty=null) => {
  const val = qty ?? document.getElementById(`buy-${sheet}-${name.replace(/\s+/g,'_')}`)?.value ?? 1;
  apiAction({ action: 'buy', sheet, name, qty: val });
};

window.addItem = () => apiAction({ action: 'add_item', item: {
  Objet: document.getElementById('n').value,
  'Prix unitaire (en crédit)': document.getElementById('pu').value,
  'poid unitaire (kg)': document.getElementById('wu').value,
  Quantité: document.getElementById('q').value,
  description: document.getElementById('d').value,
  'Range (ft)': document.getElementById('range').value,
  Hit: document.getElementById('hit').value,
  Damage: document.getElementById('damage').value,
  'bonus Armor class': document.getElementById('ac').value,
  'effet(optionel)': document.getElementById('effet').value,
}});

window.openInventoryModal = (id) => {
  const it = getInventoryItem(id);
  if (!it) return;
  modalInventoryId = id;
  modalShopRef = null;
  const entries = [
    ['Description', it.description], ['Effet', it['effet(optionel)']], ['Poids unitaire', it['poid unitaire (kg)']],
    ['Valeur unitaire', it['Prix unitaire (en crédit)']], ['Quantité', it['Quantité']], ['Range', it['Range (ft)']],
    ['Hit', it.Hit], ['Damage', it.Damage], ['Armor class', it['bonus Armor class']]
  ].filter(([,v]) => isFilled(v));
  const details = entries.map(([k,v]) => `<li><strong>${k}:</strong> ${v}</li>`).join('') || '<li>Aucune donnée.</li>';
  document.getElementById('modal-title').textContent = it.Objet || 'Objet';
  document.getElementById('modal-body').innerHTML = `
    <ul>${details}</ul>
    <div class='row'>
      <label>Qté vente <input id='sell-modal-qty' type='number' min='1' max='${it['Quantité'] || 1}' value='1' style='width:80px'></label>
      <button onclick='sellFromModal()'>Vendre</button>
      <button onclick="assignFromModal('arme')">Assigner arme</button>
      <button onclick="assignFromModal('equipement')">Assigner équipement</button>
      ${(it.type === 'arme' || it.type === 'equipement') ? `<button onclick="toggleEquipFromModal()">${it.equiped === '1' ? 'Déséquiper' : 'Équiper'}</button>` : ''}
      <button onclick='editFromModal()'>Modifier</button>
      <button onclick='closeModal()'>Fermer</button>
    </div>`;
  document.getElementById('modal-actions').innerHTML = '';
  document.getElementById('item-modal').showModal();
};

window.sellFromModal = () => {
  if (!modalInventoryId) return;
  const qty = document.getElementById('sell-modal-qty')?.value || 1;
  apiAction({ action: 'sell', id: modalInventoryId, qty });
  closeModal();
};

window.assignFromModal = (type) => {
  if (!modalInventoryId) return;
  if (type === 'arme') {
    const range = prompt('Range ?', '30');
    const hit = prompt('Hit ?', '2');
    const damage = prompt('Damage ?', '1d6');
    apiAction({ action: 'assign_type', id: modalInventoryId, type, 'Range (ft)': range, Hit: hit, Damage: damage });
    return;
  }
  const ac = prompt('Armor class ?', '1');
  const effet = prompt('Effet ?', "ho le nul il a pas d'effets");
  apiAction({ action: 'assign_type', id: modalInventoryId, type, 'bonus Armor class': ac, 'effet(optionel)': effet });
};

window.toggleEquipFromModal = () => {
  const it = getInventoryItem(modalInventoryId);
  if (!it) return;
  apiAction({ action: 'toggle_equip', id: modalInventoryId, equiped: it.equiped !== '1' });
};

window.editFromModal = () => {
  const it = getInventoryItem(modalInventoryId);
  if (!it) return;
  const fields = [
    ['description', 'Description'], ['effet(optionel)', 'Effet'], ['poid unitaire (kg)', 'Poids unitaire'],
    ['Prix unitaire (en crédit)', 'Valeur unitaire'], ['Quantité', 'Quantité'], ['Range (ft)', 'Range'],
    ['Hit', 'Hit'], ['Damage', 'Damage'], ['bonus Armor class', 'Armor class']
  ];
  const payload = { action: 'update_item', id: modalInventoryId };
  fields.forEach(([k, label]) => {
    const val = prompt(label, clean(it[k]));
    if (val !== null) payload[k] = val;
  });
  apiAction(payload);
};

window.openShopModal = (sheet, encodedName) => {
  const name = decodeURIComponent(encodedName);
  const item = (state.shop[sheet] || []).find(x => (x["nom de l'objet"] || '') === name);
  if (!item) return;
  modalShopRef = { sheet, name, price: Number(item['prix unitaire (crédit)'] || 0) };
  modalInventoryId = null;

  const imgVal = clean(item.image);
  const imgSrc = (!imgVal || imgVal === '#VALUE!') ? '' : (imgVal.startsWith('http') ? imgVal : (imgVal.includes('/') ? '/' + imgVal.replace(/^\/+/, '') : '/image/' + imgVal));
  const imgHtml = imgSrc ? `<img class='shop-preview' src='${imgSrc}' alt='${name}'>` : '';
  const details = [
    ['Description', item.description], ['Effet', item.effet], ['Poids unitaire', item['poid unitaire(kg)']],
    ['Range', item['Range (ft)']], ['Hit', item.Hit], ['Damage', item.Damage], ['Armor class', item['bonus armor class']]
  ].filter(([,v]) => isFilled(v)).map(([k,v]) => `<li><strong>${k}:</strong> ${v}</li>`).join('') || '<li>Aucune donnée.</li>';

  document.getElementById('modal-title').textContent = `Magasin - ${name}`;
  document.getElementById('modal-body').innerHTML = `
    ${imgHtml}
    <ul>${details}</ul>
    <div class='row'>
      <label>Quantité <input id='shop-qty-modal' type='number' min='1' value='1' oninput='updateShopTotal()' style='width:80px'></label>
      <strong id='shop-total-modal'>Total: ${money(modalShopRef.price)}</strong>
      <button onclick='buyFromShopModal()'>Acheter</button>
      <button onclick='closeModal()'>Fermer</button>
    </div>`;
  document.getElementById('modal-actions').innerHTML = '';
  document.getElementById('item-modal').showModal();
};

window.updateShopTotal = () => {
  if (!modalShopRef) return;
  const q = Math.max(1, Number(document.getElementById('shop-qty-modal')?.value || 1));
  document.getElementById('shop-total-modal').textContent = `Total: ${money(q * modalShopRef.price)}`;
};

window.buyFromShopModal = () => {
  if (!modalShopRef) return;
  const q = document.getElementById('shop-qty-modal')?.value || 1;
  buy(modalShopRef.sheet, modalShopRef.name, q);
};

window.showAlertModal = (message) => {
  modalInventoryId = null;
  modalShopRef = null;
  document.getElementById('modal-title').textContent = 'Alerte';
  document.getElementById('modal-body').innerHTML = `<p>${message}</p>`;
  document.getElementById('modal-actions').innerHTML = `<button onclick='closeModal()'>Fermer</button>`;
  document.getElementById('item-modal').showModal();
};

window.closeModal = () => document.getElementById('item-modal').close();

init();
