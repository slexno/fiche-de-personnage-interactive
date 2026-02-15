let state = null;

const apiAction = async (payload) => {
  const res = await fetch('/api/action', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });
  const json = await res.json();
  state = json.state;
  render();
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

const money = (v) => Number(v || 0).toFixed(2);

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
        <div class="table-wrap">
          <table><tr><th>Nom</th><th>Score</th><th>Bonus</th></tr>
            ${s.stats.map(st => `
              <tr>
                <td>${st.name}</td>
                <td><input type="number" min="1" max="20" value="${st.score}" onchange="updateStat('${st.name}', this.value)"/></td>
                <td>${st.bonus >= 0 ? '+' : ''}${st.bonus}</td>
              </tr>`).join('')}
          </table>
        </div>
        <p><strong>Armor class:</strong> ${s.armor_class}</p>
      </div>
      <div class="panel">
        <h2>Compétences</h2>
        <div class="table-wrap">
          <table><tr><th>Compétence</th><th>Modif</th><th>Bonus</th><th>Spécialisation</th></tr>
            ${s.skills.map(sk => `
              <tr>
                <td>${sk.name}</td>
                <td>${sk.mod}</td>
                <td>${sk.bonus >= 0 ? '+' : ''}${sk.bonus}</td>
                <td><input type="checkbox" ${sk.specialized ? 'checked' : ''} onchange="toggleSkill('${sk.name}', this.checked)"/></td>
              </tr>`).join('')}
          </table>
        </div>
      </div>
    </div>`;
};

const itemRow = (it, from) => {
  const to = from === 'sac à dos' ? 'coffre' : 'sac à dos';
  const maxQty = Number(it['Quantité'] || 1);
  return `
    <tr class="clickable" onclick="showItem('${it.id}')">
      <td>${it.Objet || ''}</td>
      <td>${money(it['Valeur (en crédit)'])}</td>
      <td>${money(it['Poid (kg)'])}</td>
      <td>${it['Quantité'] || 1}</td>
      <td>${it.type || 'item'}</td>
      <td>${it.equiped === '1' ? 'oui' : 'non'}</td>
      <td>
        <input id="move-${it.id}" type="number" min="1" max="${maxQty}" value="1" style="width:80px" onclick="event.stopPropagation()"/>
        <button onclick="event.stopPropagation(); transferQty('${from}', '${to}', '${it.id}')">Transférer</button>
      </td>
    </tr>`;
};

const renderInventory = () => {
  const inv = state.inventory;
  document.getElementById('inventory').innerHTML = `
    <div class="grid">
      <div class="panel">
        <h2>Sac à dos</h2>
        <div class="row">
          <strong>Crédits:</strong> ${money(inv.credits)}
          <strong>Poids:</strong> ${money(inv.bag_weight)}kg
          ${inv.overweight ? '<span class="warn">⚠ surcharge : -1 dextérité</span>' : ''}
        </div>
        <div class="row">
          <button onclick="sortBag('alpha')">Trier A-Z</button>
          <button onclick="sortBag('prix')">Trier prix</button>
          <button onclick="sortBag('poids')">Trier poids</button>
        </div>
        <div class="table-wrap">
          <table><tr><th>Objet</th><th>Valeur</th><th>Poids</th><th>Qté</th><th>Type</th><th>Équipé</th><th>Transfert</th></tr>${inv.bag.map(i => itemRow(i, 'sac à dos')).join('')}</table>
        </div>

        <h3>Ajouter manuellement</h3>
        <div class="row">
          <input id="n" placeholder="Nom"/>
          <input id="pu" type="number" step="0.1" placeholder="Prix unitaire"/>
          <input id="wu" type="number" step="0.1" placeholder="Poids unitaire"/>
          <input id="q" type="number" value="1"/>
          <input id="d" placeholder="Description" style="min-width:220px"/>
          <button onclick="addItem()">Ajouter</button>
        </div>
      </div>

      <div class="panel">
        <h2>Coffre</h2>
        <div class="table-wrap">
          <table><tr><th>Objet</th><th>Valeur</th><th>Poids</th><th>Qté</th><th>Type</th><th>Équipé</th><th>Transfert</th></tr>${inv.chest.map(i => itemRow(i, 'coffre')).join('')}</table>
        </div>

        <h3>Armes équipables (max 4)</h3>
        <ul>${inv.weapons.map(w => `<li>${w.Objet} (${w['Range (ft)'] || '-'}ft, ${w.Hit || '-'}, ${w.Damage || '-'}) <button onclick="equip('${w.id}', ${w.equiped !== '1'})">${w.equiped === '1' ? 'Déséquiper' : 'Équiper'}</button></li>`).join('')}</ul>

        <h3>Équipements équipables (max 3)</h3>
        <ul>${inv.equipments.map(e => `<li>${e.Objet} (AC +${e['bonus Armor class'] || 0}) <button onclick="equip('${e.id}', ${e.equiped !== '1'})">${e.equiped === '1' ? 'Déséquiper' : 'Équiper'}</button></li>`).join('')}</ul>

        <h3>Vendre / Assigner</h3>
        <div class="row">
          <input id="sell-id" placeholder="ID objet"/>
          <input id="sell-q" type="number" value="1"/>
          <button onclick="sell()">Vendre</button>
          <button onclick="assign('arme')">Assigner arme</button>
          <button onclick="assign('equipement')">Assigner équipement</button>
        </div>
        <p class="small">Cliquez un objet pour ouvrir une popup de détails et remplir automatiquement l'ID.</p>
      </div>
    </div>`;
};

const renderShop = () => {
  const shop = state.shop;
  const sections = Object.entries(shop).map(([sheet, items]) => `
    <div class="panel">
      <h3>${sheet}</h3>
      <div class="table-wrap">
        <table><tr><th>Objet</th><th>Prix</th><th>Poids</th><th>Description</th><th>Achat</th></tr>
          ${items.map(i => {
            const name = i["nom de l'objet"] || '';
            const inputId = `buy-${sheet}-${name.replace(/\s+/g, '_')}`;
            return `<tr>
              <td>${name}</td>
              <td>${money(i['prix unitaire (crédit)'])}</td>
              <td>${money(i['poid unitaire(kg)'])}</td>
              <td>${i.description || ''}</td>
              <td><input id="${inputId}" type="number" value="1" style="width:80px"/><button onclick="buyEncoded('${sheet}','${encodeURIComponent(name)}')" style="margin-left:6px">Acheter</button></td>
            </tr>`;
          }).join('')}
        </table>
      </div>
    </div>`).join('');
  document.getElementById('shop').innerHTML = sections;
};

window.updateStat = (name, score) => apiAction({ action: 'update_stat', name, score });
window.toggleSkill = (name, specialized) => apiAction({ action: 'toggle_skill', name, specialized });
window.sortBag = (key) => apiAction({ action: 'sort', key, source: 'sac à dos' });
window.equip = (id, equiped) => apiAction({ action: 'toggle_equip', id, equiped });
window.sell = () => apiAction({ action: 'sell', id: document.getElementById('sell-id').value, qty: document.getElementById('sell-q').value });
window.buy = (sheet, name) => {
  const id = `buy-${sheet}-${name.replace(/\s+/g, '_')}`;
  apiAction({ action: 'buy', sheet, name, qty: document.getElementById(id).value });
};
window.buyEncoded = (sheet, encodedName) => {
  const name = decodeURIComponent(encodedName);
  window.buy(sheet, name);
};
window.addItem = () => apiAction({
  action: 'add_item',
  item: {
    Objet: document.getElementById('n').value,
    'Prix unitaire (en crédit)': document.getElementById('pu').value,
    'poid unitaire (kg)': document.getElementById('wu').value,
    Quantité: document.getElementById('q').value,
    description: document.getElementById('d').value,
  }
});
window.transferQty = (from, to, id) => {
  const qty = document.getElementById(`move-${id}`)?.value || 1;
  apiAction({ action: 'transfer_item', from, to, id, qty });
};

window.showItem = (id) => {
  const all = [...state.inventory.bag, ...state.inventory.chest];
  const it = all.find(x => x.id === id);
  if (!it) return;
  document.getElementById('sell-id').value = id;
  document.getElementById('modal-title').textContent = it.Objet || 'Objet';
  const desc = [
    `ID: ${it.id}`,
    `Type: ${it.type || 'item'}`,
    `Quantité: ${it['Quantité'] || 1}`,
    `Prix unitaire: ${money(it['Prix unitaire (en crédit)'])}`,
    `Poids unitaire: ${money(it['poid unitaire (kg)'])} kg`,
    `Description: ${it.description || 'Aucune description'}`,
    it['Range (ft)'] ? `Portée: ${it['Range (ft)']} ft` : '',
    it.Hit ? `Hit: ${it.Hit}` : '',
    it.Damage ? `Damage: ${it.Damage}` : '',
    it['bonus Armor class'] ? `Bonus AC: ${it['bonus Armor class']}` : '',
    it['effet(optionel)'] ? `Effet: ${it['effet(optionel)']}` : ''
  ].filter(Boolean).join('\n');
  document.getElementById('modal-body').textContent = desc;
  document.getElementById('item-modal').showModal();
};

window.closeModal = () => document.getElementById('item-modal').close();

window.assign = (type) => {
  const id = document.getElementById('sell-id').value;
  if (!id) return;
  if (type === 'arme') {
    const range = prompt('Range (ft) ?', '30');
    const hit = prompt('Hit ?', '2');
    const damage = prompt('Damage ?', '1d6');
    apiAction({ action: 'assign_type', id, type, 'Range (ft)': range, Hit: hit, Damage: damage });
    return;
  }
  const ac = prompt('Bonus Armor class ?', '1');
  const effet = prompt('Effet optionnel ?', "ho le nul il a pas d'effets");
  apiAction({ action: 'assign_type', id, type, 'bonus Armor class': ac, 'effet(optionel)': effet });
};

init();
