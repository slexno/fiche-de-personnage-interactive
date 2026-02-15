let state = null;

async function api(path, payload) {
  const res = await fetch(path, {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)});
  const data = await res.json();
  if (!res.ok) alert(data.message || 'Erreur');
  if (data.state) {
    state = data.state;
    render();
  }
}

async function load() {
  state = await (await fetch('/api/state')).json();
  render();
}

function render() { renderStats(); renderInventory(); renderShop(); }

function renderStats() {
  const root = document.getElementById('stats');
  root.innerHTML = `<h2>Statistiques</h2>
    <p>Armor Class: <b>${state.armor_class}</b></p>
    <div class="grid"><div class="card"><h3>Scores</h3><table><tr><th>Stat</th><th>Score</th><th>Bonus</th></tr>
    ${state.stats.map(s => `<tr><td>${s.name}</td><td><input type='number' min='0' max='20' value='${s.score}' onchange="api('/api/stat',{name:'${s.name}',value:this.value})"></td><td>${s.bonus}</td></tr>`).join('')}
    </table></div>
    <div class="card"><h3>Compétences</h3><table><tr><th>Nom</th><th>Modif</th><th>Bonus</th><th>Spécialisé</th></tr>
    ${state.skills.map(s => `<tr><td>${s.name}</td><td>${s.modifier}</td><td>${s.bonus}</td><td><input type='checkbox' ${s.specialized?'checked':''} onchange="api('/api/skill',{name:'${s.name}',specialized:this.checked})"></td></tr>`).join('')}
    </table></div></div>`;
}

function itemTable(items, source) {
  return `<table><tr><th>Objet</th><th>Valeur totale</th><th>Poids total</th><th>Quantité</th><th>Action</th></tr>
    ${items.map(i => `<tr><td title='${i.description || ''}'>${i.name}</td><td>${i.total_value}</td><td>${i.total_weight}</td><td>${i.qty}</td>
    <td><button onclick="const q=prompt('Quantité?','1');if(q)api('/api/transfer',{source:'${source}',name:'${i.name}',qty:Number(q)})">Transférer</button></td></tr>`).join('')}</table>`;
}

function renderInventory() {
  const root = document.getElementById('inventory');
  const warning = state.weight_warning ? `<span class='warning'>⚠ Sac > 50kg : -1 Dextérité</span>` : '';
  root.innerHTML = `<h2>Inventaire</h2><p>Crédits: <b>${state.credits}</b> ${warning}</p>
    <div class='grid'><div class='card'><h3>Sac à dos</h3>${itemTable(state.bag,'bag')}</div>
    <div class='card'><h3>Coffre</h3>${itemTable(state.chest,'chest')}</div></div>
    <div class='card'>
      <h3>Ajouter un objet</h3>
      <input id='new-name' placeholder='nom'> <input id='new-weight' placeholder='poids unitaire'> <input id='new-price' placeholder='prix unitaire'> <input id='new-qty' placeholder='quantité'>
      <button onclick="api('/api/item',{target:'bag',name:gid('new-name').value,unit_weight:Number(gid('new-weight').value),unit_price:Number(gid('new-price').value),qty:Number(gid('new-qty').value),description:''})">Ajouter au sac</button>
    </div>
    <div class='grid'><div class='card'><h3>Armes (max 4 équipées)</h3><table><tr><th>Nom</th><th>Range</th><th>Hit</th><th>Damage</th><th>Eq</th></tr>
      ${state.weapons.map(w => `<tr><td>${w.name}</td><td>${w.range}</td><td>${w.hit}</td><td>${w.damage}</td><td><input type='checkbox' ${w.equipped?'checked':''} onchange="api('/api/equip',{kind:'weapon',name:'${w.name}',equipped:this.checked})"></td></tr>`).join('')}
    </table></div>
    <div class='card'><h3>Equipements (max 3 équipés)</h3><table><tr><th>Nom</th><th>AC bonus</th><th>Effet</th><th>Eq</th></tr>
      ${state.equipments.map(e => `<tr><td>${e.name}</td><td>${e.ac_bonus}</td><td>${e.effect || "ho le nul il a pas d'effets"}</td><td><input type='checkbox' ${e.equipped?'checked':''} onchange="api('/api/equip',{kind:'equipment',name:'${e.name}',equipped:this.checked})"></td></tr>`).join('')}
    </table></div></div>`;
}

function renderShop() {
  const root = document.getElementById('shop');
  const cats = Object.keys(state.shop);
  root.innerHTML = `<h2>Magasin</h2>${cats.map(c => `<div class='card'><h3>${c}</h3><table><tr><th>Objet</th><th>Prix</th><th>Description</th><th>Poids</th><th>Action</th></tr>
    ${state.shop[c].map(i => `<tr><td>${i["nom de l'objet"] || ''}</td><td>${i['prix unitaire (crédit)'] || ''}</td><td>${i.description || ''}</td><td>${i['poid unitaire(kg)'] || ''}</td><td><button onclick="const q=prompt('Quantité','1');if(q)api('/api/buy',{category:'${c}',name:${JSON.stringify(i["nom de l'objet"] || '')},qty:Number(q)})">Acheter</button></td></tr>`).join('')}
  </table></div>`).join('')}
  <div class='card'><h3>Vendre</h3><p>Depuis sac/coffre, choisis objet et quantité.</p>
  <input id='sell-src' placeholder='bag ou chest' value='bag'> <input id='sell-name' placeholder='nom objet'> <input id='sell-qty' placeholder='quantité'>
  <button onclick="api('/api/sell',{source:gid('sell-src').value,name:gid('sell-name').value,qty:Number(gid('sell-qty').value)})">Vendre</button>
  </div>`;
}

function gid(id){ return document.getElementById(id); }

document.querySelectorAll('.tabs button').forEach(btn => btn.addEventListener('click', () => {
  document.querySelectorAll('.tab').forEach(el => el.classList.add('hidden'));
  gid(btn.dataset.tab).classList.remove('hidden');
}));

load();
