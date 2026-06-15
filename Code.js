// ============================================================
// MithaiCalc v13.0 — MERGED [Cost Calculator + Rate Manager]
// ============================================================
// HTML Files Required:
//   Index.html   → MithaiCalc Cost Calculator (Admin)
//   Staff.html   → MithaiCalc Cost Calculator (Staff)
//   NewIndex.html → Rate Manager (Raw Material Rate Update)
//
// doGet routing:
//   ?role=staff   → Staff.html
//   ?role=rates   → NewIndex.html  (Rate Manager)
//   (default)     → Index.html     (Admin Cost Calculator)
// ============================================================
 
// ── RM_Price flat sheet config ──
var RM_SHEET      = 'RM_Price';
var RM_COL_ENG    = 1;  // A
var RM_COL_HIN    = 2;  // B
var RM_COL_CAT    = 3;  // C
var RM_COL_UNIT   = 4;  // D
var RM_COL_PRICE  = 5;  // E
var RM_HDR        = ['Material_Name_Eng','Material_Name_Hindi','Category','Unit','Price_Per_Unit'];
var RM_DATA_START = 2;
 
// ── Rate Manager Config (from NewCode.gs) ──
var SPREADSHEET_ID = '16kP9AjSCXE1nGXbezaMNHM-WFM7ywR0JdpMyj77rHhI';
var DEPT_SHEETS    = ['Nasta Raw', 'Savli Bakery', 'Namkeen Raw', 'Mithai Raw'];
 
// ── Dashboard Category -> Raw Sheet Mapping Config ──
var CAT_SHEET_CONFIG = {
  // flat:false → dual-block layout: A=Hindi, B=Eng, C=Rate | G=Eng, H=Rate
  // flat:true  → flat 5-col layout: A=Eng, B=Hindi, C=Category, D=Unit, E=Price_Per_Unit
  'Nasta Raw'   : { cat:'Nastha',  flat:false, sets:[{engCol:2,hinCol:1,rateCol:3},{engCol:7,hinCol:null,rateCol:8}] },
  'Namkeen Raw' : { cat:'Namkeen', flat:false, sets:[{engCol:2,hinCol:1,rateCol:3},{engCol:7,hinCol:null,rateCol:8}] },
  'Savli Bakery': { cat:'Bakery',  flat:false, sets:[{engCol:2,hinCol:1,rateCol:3},{engCol:7,hinCol:null,rateCol:8}] },
  'Mithai Raw'  : { cat:'Mithai',  flat:true,  sets:[{engCol:1,hinCol:2,rateCol:5}] }
};
 
// UI Dashboard category names mapped directly to the destination Sheets
var DASHBOARD_SHEET_MAP = {
  'nastha'  : 'Nasta Raw',
  'nasta'   : 'Nasta Raw',
  'bakery'  : 'Savli Bakery',
  'namkeen' : 'Namkeen Raw',
  'mithai'  : 'Mithai Raw',
  'other'   : 'Mithai Raw',   // fallback to Mithai Raw since New Added sheet removed
  'sub-rm'  : 'Mithai Raw'    // fallback to Mithai Raw since New Added sheet removed
};
 
var CAT_BLOCKS  = ['Nastha','Bakery','Namkeen','Mithai'];
var ALL_CATS    = ['Nastha','Bakery','Namkeen','Mithai','Other','Sub-RM'];
var GST         = {'Nastha':5,'Bakery':5,'Namkeen':5,'Mithai':5};
var FG_CATS     = ['Nastha','Bakery','Namkeen','Mithai'];
var OH_SUB_CATS = ['Nastha','Bakery','Namkeen','Mithai'];
 
// ════════════════════════════════════════════════════════════════════
// ── UTILITIES ──
// ════════════════════════════════════════════════════════════════════
function pn(v){
  if(v===null||v===undefined||v==='')return 0;
  if(typeof v==='number')return isNaN(v)?0:v;
  var n=parseFloat(String(v).replace(/,/g,'').replace(/%/g,'').trim());
  return isNaN(n)?0:n;
}
function trim(v){return String(v===null||v===undefined?'':v).trim();}
function kw(s){return trim(s).toLowerCase().replace(/[^a-z0-9]/g,'');}
function ckGS(c){return String(c||'').toLowerCase().replace(/[^a-z0-9]/g,'_').replace(/_+/g,'_').replace(/^_|_$/g,'');}
function SS(){return SpreadsheetApp.getActiveSpreadsheet();}
function sh(n){var s=SS().getSheetByName(n);if(!s)throw new Error('Sheet "'+n+'" not found');return s;}
 
function ensureSh(n,headers){
  var ss=SS(),s=ss.getSheetByName(n);
  if(!s){
    s=ss.insertSheet(n);
    if(headers&&headers.length){
      s.appendRow(headers);
      s.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#3B1F0E').setFontColor('#FFF');
      s.setFrozenRows(1);
    }
  }
  return s;
}
 
function rd(name){
  var s=sh(name),data=s.getDataRange().getValues();
  if(data.length<2)return{h:[],rows:[],ds:1,s:s,raw:data};
  var hr=0,KEYS=['fgname','material','ingredient','category','unit','price','percentage','qty','overhead','cost','gst','subrmname'];
  for(var i=0;i<Math.min(5,data.length);i++){
    var score=0;
    data[i].forEach(function(c){var k=kw(c);KEYS.forEach(function(x){if(k.indexOf(x)>-1||x.indexOf(k)>-1)score++;});});
    if(score>=2){hr=i;break;}
  }
  var h=data[hr].map(function(x){return trim(x);});
  var rows=data.slice(hr+1).filter(function(r){return trim(r[0])&&!(/[^\x00-\x7F]/.test(trim(r[0]))||trim(r[0])==='');}).map(function(r){
    var o={};h.forEach(function(k,i){o[k]=r[i]!==undefined?r[i]:'';});return o;
  });
  return{h:h,rows:rows,ds:hr+1,s:s,raw:data};
}
 
function ci(headers,keywords){
  var ks=headers.map(kw);
  for(var j=0;j<keywords.length;j++)for(var i=0;i<ks.length;i++)if(ks[i]===keywords[j]||ks[i].indexOf(keywords[j])>-1||keywords[j].indexOf(ks[i])>-1)return i;
  return -1;
}
 
// ════════════════════════════════════════════════════════════════════
// RM_PRICE FLAT SHEET ENGINE
// ════════════════════════════════════════════════════════════════════
function ensureRMSheet(){
  var ss=SS(),s=ss.getSheetByName(RM_SHEET);
  if(!s){
    s=ss.insertSheet(RM_SHEET);
    s.appendRow(RM_HDR);
    s.getRange(1,1,1,RM_HDR.length).setFontWeight('bold').setBackground('#3B1F0E').setFontColor('#FFF');
    s.setFrozenRows(1);
  }
  return s;
}
 
function _readRMAll(){
  var s=ensureRMSheet(),last=s.getLastRow();
  if(last<RM_DATA_START)return [];
  var vals=s.getRange(RM_DATA_START,1,last-RM_DATA_START+1,5).getValues();
  var out=[];
  vals.forEach(function(r,i){
    var name=trim(String(r[0]||''));
    if(!name)return;
    if(/[^\x00-\x7F]/.test(name))return;
    out.push({
      name    :name,
      nameHi  :trim(String(r[1]||'')),
      category:trim(String(r[2]||'')),
      unit    :trim(String(r[3]||''))||'Kg',
      price   :pn(r[4]),
      rowNum  :RM_DATA_START+i
    });
  });
  return out;
}
 
function readAllRM(){
  return _readRMAll().map(function(r){
    return{name:r.name,nameHi:r.nameHi,category:r.category,unit:r.unit,price:r.price};
  });
}
 
function _findRowInRMWithCat(engName, category){
  var s=ensureRMSheet(),last=s.getLastRow();
  if(last<RM_DATA_START)return -1;
  var kwN=kw(engName);
  var kwC=kw(category);
  var vals=s.getRange(RM_DATA_START,1,last-RM_DATA_START+1,3).getValues();
  for(var i=0;i<vals.length;i++){
    if(kw(trim(String(vals[i][0]||'')))===kwN && kw(trim(String(vals[i][2]||'')))===kwC) {
      return RM_DATA_START+i;
    }
  }
  return -1;
}
 
function _buildPriceMap(){
  var pm={};
  _readRMAll().forEach(function(r){
    if(!r.name)return;
    var e={price:r.price,unit:r.unit,category:r.category};
    var nameKw = kw(r.name);
    var catKw = kw(r.category);
 
    pm[nameKw + '_' + catKw] = e;
    pm[r.name.toLowerCase().trim() + '_' + r.category.toLowerCase().trim()] = e;
 
    if(!pm[nameKw]) {
      pm[nameKw] = e;
    }
  });
  return pm;
}
 
// ── RAW DATA INGESTION — flat (Mithai Raw) and dual-block (Nasta/Namkeen/Bakery) ──
function _readCatSheetItems(sheetName, cfg){
  var ss=SS(), s=ss.getSheetByName(sheetName);
  if(!s)return[];
  var last=s.getLastRow();
  if(last<2)return[];
  var items=[], seen={};
 
  if(cfg.flat){
    // ── FLAT 5-column (Mithai Raw, RM_Price) ──
    // A=Material_Name_Eng | B=Material_Name_Hindi | C=Category | D=Unit | E=Price_Per_Unit
    // Header = row 1 → read from row 2 onward
    var vals=s.getRange(2,1,last-1,5).getValues();
    for(var i=0;i<vals.length;i++){
      var row=vals[i];
      var name=trim(String(row[0]||''));
      if(!name)continue;
      if(/[^-]/.test(name))continue;
      var key=kw(name)+'_'+kw(cfg.cat);
      if(!seen[key]){
        seen[key]=true;
        items.push({
          name    :name,
          nameHi  :trim(String(row[1]||'')),
          rate    :pn(row[4]),
          category:cfg.cat,
          unit    :trim(String(row[3]||''))||'Kg'
        });
      }
    }
  } else {
    // ── DUAL-BLOCK (Nasta Raw, Namkeen Raw, Savli Bakery) ──
    // Block 1: A=Hindi(1), B=Eng(2), C=Rate(3)
    // Block 2: G=Eng(7),   H=Rate(8)
    var vals=s.getRange(1,1,last,9).getValues();
    cfg.sets.forEach(function(set){
      var ec=set.engCol-1, hc=set.hinCol!=null?set.hinCol-1:-1, rc=set.rateCol-1;
      for(var i=1;i<vals.length;i++){   // i=1 skips header row
        var row=vals[i];
        var name=trim(String(row[ec]||''));
        if(!name)continue;
        if(/[^-]/.test(name))continue;
        var key=kw(name)+'_'+kw(cfg.cat);
        if(!seen[key]){
          seen[key]=true;
          items.push({
            name    :name,
            nameHi  :hc>=0?trim(String(row[hc]||'')):'',
            rate    :pn(row[rc]),
            category:cfg.cat,
            unit    :'Kg'
          });
        }
      }
    });
  }
  return items;
}
 
// _readNewAddedItems() removed — 'New Added' sheet no longer exists
 
// ════════════════════════════════════════════════════════════════════
// ⚡ updateAllRecipeRates — HIGH-SPEED BATCH ENGINE
// ════════════════════════════════════════════════════════════════════
function updateAllRecipeRates(){
  try{
    var rmSh = ensureRMSheet();
    var rmDataRange = rmSh.getDataRange();
    var rmValues = rmDataRange.getValues();
 
    var rmHeaders = RM_HDR;
    var existingRows = rmValues.length > 1 ? rmValues.slice(1) : [];
 
    existingRows = existingRows.map(function(r) {
      return [
        r[0] !== undefined ? r[0] : '',
        r[1] !== undefined ? r[1] : '',
        r[2] !== undefined ? r[2] : '',
        r[3] !== undefined ? r[3] : '',
        r[4] !== undefined ? r[4] : 0
      ];
    });
 
    var lookupMap = {};
    existingRows.forEach(function(r, idx) {
      var name = trim(r[0]);
      var cat = trim(r[2]);
      if(name) {
        lookupMap[kw(name) + '_' + kw(cat)] = idx;
      }
    });
 
    var allItems = [];
    Object.keys(CAT_SHEET_CONFIG).forEach(function(sheetName){
      var cfg = CAT_SHEET_CONFIG[sheetName];
      _readCatSheetItems(sheetName,cfg).forEach(function(item){ allItems.push(item); });
    });
    // _readNewAddedItems: 'New Added' sheet has been removed
 
    var updated = 0, added = 0;
    var seenKeys = {};
 
    allItems.forEach(function(item){
      if(!item.name) return;
      var compositeKey = kw(item.name) + '_' + kw(item.category);
      if(seenKeys[compositeKey]) return;
      seenKeys[compositeKey] = true;
 
      if(lookupMap[compositeKey] !== undefined) {
        var existingIdx = lookupMap[compositeKey];
        existingRows[existingIdx][4] = item.rate;
        updated++;
      } else {
        var newRow = [item.name, item.nameHi || '', item.category || 'Other', item.unit || 'Kg', item.rate];
        existingRows.push(newRow);
        lookupMap[compositeKey] = existingRows.length - 1;
        added++;
      }
    });
 
    var pm = {};
    existingRows.forEach(function(r) {
      var name = trim(r[0]);
      var cat = trim(r[2]);
      if(!name) return;
      var e = { price: pn(r[4]), unit: trim(r[3]), category: cat };
      pm[kw(name) + '_' + kw(cat)] = e;
      if (!pm[kw(name)]) pm[kw(name)] = e;
    });
 
    var subMap = buildSubRmMap();
    var subUpdated = 0;
 
    Object.keys(subMap).forEach(function(sn){
      var s = subMap[sn], totalCost = 0, oq = pn(s.outputQty) || 1;
      s.ingredients.forEach(function(g){
        var entry = pm[kw(g.material) + '_subrm'] || pm[kw(g.material)] || {price:0};
        totalCost += pn(g.qty) * entry.price;
      });
      var cpu = totalCost / oq;
      if(cpu > 0){
        var compositeKey = kw(s.name) + '_subrm';
        if(lookupMap[compositeKey] !== undefined) {
          var existingIdx = lookupMap[compositeKey];
          existingRows[existingIdx][4] = cpu;
        } else {
          var newSubRow = [s.name, '', 'Sub-RM', s.outputUnit || 'Kg', cpu];
          existingRows.push(newSubRow);
          lookupMap[compositeKey] = existingRows.length - 1;
        }
        subUpdated++;
      }
    });
 
    var finalOutputValues = [rmHeaders].concat(existingRows);
    rmSh.clearContents();
    rmSh.getRange(1, 1, finalOutputValues.length, 5).setValues(finalOutputValues);
    // Header row: bold white on dark brown
    rmSh.getRange(1, 1, 1, 5).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#3B1F0E');
    // All data rows: bold black font, no background
    if(existingRows.length > 0){
      rmSh.getRange(2, 1, existingRows.length, 5)
          .setFontWeight('bold').setFontColor('#000000').setBackground(null);
    }
 
    // ── Recipes sync: update Unit_Rate + Line_Total for every material whose rate changed ──
    // Build rate map from final existingRows: kw(name) → new price
    var recipeRateMap = {};
    existingRows.forEach(function(r){
      var n = trim(String(r[0]||''));
      if(n) recipeRateMap[kw(n)] = pn(r[4]);
    });
    var recSyncCount = 0;
    try {
      var recSh2 = SS().getSheetByName('Recipes');
      if(recSh2 && recSh2.getLastRow() >= 2){
        // Cols: A=FG_Name(0), B=Material_Name(1), C=Qty(2), D=Unit(3), E=Unit_Rate(4), F=Line_Total(5)
        var recRng2  = recSh2.getRange(2, 1, recSh2.getLastRow()-1, 6);
        var recVals2 = recRng2.getValues();
        for(var ri=0; ri<recVals2.length; ri++){
          var mat = trim(String(recVals2[ri][1]||''));
          if(!mat) continue;
          var nr = recipeRateMap[kw(mat)];
          if(nr === undefined) continue;
          recVals2[ri][4] = nr;                        // Unit_Rate
          recVals2[ri][5] = pn(recVals2[ri][2]) * nr; // Line_Total = Qty × Rate
          recSyncCount++;
        }
        if(recSyncCount > 0) recRng2.setValues(recVals2);
      }
    } catch(re){ Logger.log('Recipe sync in updateAllRecipeRates: '+re.message); }
 
    SpreadsheetApp.flush();
 
    var freshRM = readAllRM();
    var parts = [];
    if(updated > 0)      parts.push('Updated '  + updated      + ' rates');
    if(added > 0)        parts.push('Added '     + added        + ' new items');
    if(subUpdated > 0)   parts.push('Sub-RM '    + subUpdated   + ' recalculated');
    if(recSyncCount > 0) parts.push('Recipes '   + recSyncCount + ' lines synced');
    var msg = '✅ ' + (parts.length ? parts.join(' | ') : 'No changes found') + ' (Sync Optimized)';
 
    return {success: true, message: msg, updated: updated, added: added, subUpdated: subUpdated, rm: freshRM};
  } catch(e) {
    throw new Error('Update Rates failed: ' + e.message);
  }
}
 
// ════════════════════════════════════════════════════════════════════
// RAW MATERIAL CRUD
// ════════════════════════════════════════════════════════════════════
function saveRM(row,mode){
  var eng  =trim(row.name);
  var hin  =trim(row.nameHi||'');
  var cat  =trim(row.category)||'Other';
  var unit =trim(row.unit)||'Kg';
  var price=pn(row.price);
  if(!eng)throw new Error('Material name required');
 
  if(!hin){try{hin=LanguageApp.translate(eng,'en','hi');}catch(e){hin='';}}
 
  // 1. Dynamic Routing Engine to push elements into correct category source sheets
  var targetSheetName = DASHBOARD_SHEET_MAP[kw(cat)] || 'Mithai Raw';
  var targetSh = ensureSh(targetSheetName);
 
  // All category sheets use flat 5-column format: [Eng, Hindi, Category, Unit, Price]
  targetSh.appendRow([eng, hin, cat, unit, price]);
 
  // 2. Mirror element adjustments out onto RM_Price Matrix Database
  var rmSh=ensureRMSheet();
  var existRow=_findRowInRMWithCat(eng, cat);
  if(mode==='add'&&existRow>0)throw new Error('"'+eng+'" already exists in category: '+cat);
 
  if(existRow>0){
    rmSh.getRange(existRow,RM_COL_ENG,  1,1).setValue(eng);
    rmSh.getRange(existRow,RM_COL_HIN,  1,1).setValue(hin);
    rmSh.getRange(existRow,RM_COL_CAT,  1,1).setValue(cat);
    rmSh.getRange(existRow,RM_COL_UNIT, 1,1).setValue(unit);
    rmSh.getRange(existRow,RM_COL_PRICE,1,1).setValue(price);
  }else{
    rmSh.appendRow([eng,hin,cat,unit,price]);
    // Keep new row bold + consistent with existing RM_Price rows
    var newRmRow = rmSh.getLastRow();
    rmSh.getRange(newRmRow, 1, 1, 5).setFontWeight('bold').setFontColor('#000000').setBackground(null);
  }
  SpreadsheetApp.flush();
  return{success:true,message:'"'+eng+'" saved to '+targetSheetName+' and RM_Price.',nameHi:hin};
}
 
function deleteRM(name,category){
  name=trim(name);
  var existRow=_findRowInRMWithCat(name, category);
  if(existRow<0)throw new Error('"'+name+'" not found in category: '+category);
  ensureRMSheet().deleteRow(existRow);
  SpreadsheetApp.flush();
  return{success:true,message:'"'+name+'" deleted.'};
}
 
// ════════════════════════════════════════════════════════════════════
// ⚡ SAVE RECIPE — PERFORMANCE ENGINE UPGRADE (ULTRA-FAST IN-MEMORY CLEAN)
// ════════════════════════════════════════════════════════════════════
function saveRecipe(fgName,ings,batchQty,batchUnit,replace){
  fgName=trim(fgName); batchQty=pn(batchQty)||1; batchUnit=trim(batchUnit)||'Kg';
  if(!fgName)throw new Error('Product name required');
  if(!ings||!ings.length)throw new Error('No ingredients provided');
 
  var recipeSh = ensureSh('Recipes',['FG_Name','Material_Name','Qty_Required','Unit','Unit_Rate','Line_Total','Final_Product_Qty','Final_Product_Unit']);
  var fullValues = recipeSh.getDataRange().getValues();
  var headers = fullValues[0];
  var keptRows = [];
 
  // Filter out any matching older structural rows inside memory array loop
  if(replace && fullValues.length > 1){
    var targetKw = kw(fgName);
    for(var i = 1; i < fullValues.length; i++){
      if(kw(trim(fullValues[i][0])) !== targetKw){
        keptRows.push(fullValues[i]);
      }
    }
  } else if(fullValues.length > 1) {
    keptRows = fullValues.slice(1);
  }
 
  // Append new item objects directly into active working memory block arrays
  ings.forEach(function(g, idx){
    var mat=trim(g.material), qty=pn(g.qty), unit=trim(g.unit)||'Kg', rate=pn(g.rate)||0;
    if(!mat) return;
    var lineTotal = qty * rate;
    keptRows.push([fgName, mat, qty, unit, rate, lineTotal, idx===0?batchQty:'', idx===0?batchUnit:'']);
  });
 
  // Wipe sheet out completely and overwrite with new compiled structure array in a single batch pass
  recipeSh.clearContents();
  var finalOutputMatrix = [headers].concat(keptRows);
  recipeSh.getRange(1, 1, finalOutputMatrix.length, finalOutputMatrix[0].length).setValues(finalOutputMatrix);
 
  SpreadsheetApp.flush();
  return{success:true,message:'"'+fgName+'" recipe saved successfully.'};
}
 
// ════════════════════════════════════════════════════════════════════
// MAIN SYSTEM BOOT STRAPPER
// ════════════════════════════════════════════════════════════════════
function getData(){
  ensureRMSheet();
  var rm=readAllRM();
  var fg=[],rec={},subRm={},oh=[],recRaw=[];
  try{
    fg=rd('FG_List').rows.map(function(r){
      return{name:trim(r['FG_Name']),category:trim(r['Category']),gst:pn(r['GST_Pct']),notes:trim(r['Notes'])};
    }).filter(function(r){return r.name;});
  }catch(e){Logger.log('FG:'+e);}
  try{recRaw=rd('Recipes').raw;rec=buildRecMap(recRaw,rm);}catch(e){Logger.log('REC:'+e);}
  try{subRm=buildSubRmMap();}catch(e){Logger.log('SubRM:'+e);}
  oh=_readOHRows();
  return{
    fg:fg,rm:rm,recipes:rec,subRm:subRm,oh:oh,
    ohPct:oh.reduce(function(s,r){return s+r.percentage;},0),
    gst:GST,rmCats:CAT_BLOCKS,fgCats:FG_CATS,ohSubCats:OH_SUB_CATS,
    recCount:recRaw.length,rmCount:rm.length,
    blocks:CAT_BLOCKS,
    catSheetMap:{'Nastha':'Nasta Raw','Bakery':'Savli Bakery','Namkeen':'Namkeen Raw','Mithai':'Mithai Raw','Other':'Mithai Raw'}
  };
}
 
function _readOHRows(){
  try{
    return rd('Overheads').rows
      .filter(function(r){return trim(r['Category'])&&!trim(r['Category']).toUpperCase().includes('TOTAL');})
      .map(function(r){return{
        category   :trim(r['Category']||''),
        subCategory:trim(r['Sub_Category']||''),
        basis      :trim(r['Calculation_Basis']||'% of RM Cost'),
        percentage :pn(r['Percentage']||0)
      };});
  }catch(e){return[];}
}
 
function buildRecMap(raw,rmList){
  var map={},pm={};
  if(rmList&&rmList.length)rmList.forEach(function(r){
    if(r.name){
      var nameKw = kw(r.name);
      var catKw = kw(r.category);
      pm[nameKw + '_' + catKw] = {price:r.price,unit:r.unit};
      if(!pm[nameKw]) {
        pm[nameKw] = {price:r.price,unit:r.unit};
      }
    }
  });
  if(!raw||raw.length<2)return map;
 
  var hr=0;
  for(var i=0;i<Math.min(5,raw.length);i++){
    var s=0;
    raw[i].forEach(function(c){var k=kw(c);if(['fgname','materialname','qtyrequired','unit'].some(function(x){return k.indexOf(x)>-1;}))s++;});
    if(s>=2){hr=i;break;}
  }
 
  var h=raw[hr],colIdx={};
  h.forEach(function(cell,idx){colIdx[kw(cell)]=idx;});
 
  var FG   = colIdx[kw('FG_Name')]           !== undefined ? colIdx[kw('FG_Name')]           : 0;
  var MAT  = colIdx[kw('Material_Name')]     !== undefined ? colIdx[kw('Material_Name')]     : 1;
  var QTY  = colIdx[kw('Qty_Required')]      !== undefined ? colIdx[kw('Qty_Required')]      : 2;
  var UNIT = colIdx[kw('Unit')]              !== undefined ? colIdx[kw('Unit')]              : 3;
  var RATE = colIdx[kw('Unit_Rate')]         !== undefined ? colIdx[kw('Unit_Rate')]         : 4;
  var BQTY = colIdx[kw('Final_Product_Qty')] !== undefined ? colIdx[kw('Final_Product_Qty')] : 6;
  var BUNT = colIdx[kw('Final_Product_Unit')]!== undefined ? colIdx[kw('Final_Product_Unit')]: 7;
 
  var batchInfo={};
  for(var r=hr+1;r<raw.length;r++){
    var row=raw[r],fg2=trim(row[FG]);if(!fg2)continue;
    var fgl=fg2.toLowerCase();
    if(!batchInfo[fgl])batchInfo[fgl]={qty:0,unit:''};
    var bq=pn(row[BQTY]);if(bq>0&&!batchInfo[fgl].qty)batchInfo[fgl].qty=bq;
    var bu=trim(String(row[BUNT]||''));if(bu&&!batchInfo[fgl].unit)batchInfo[fgl].unit=bu;
  }
 
  for(var r=hr+1;r<raw.length;r++){
    var row=raw[r],fg2=trim(row[FG]);if(!fg2)continue;
    var fgl=fg2.toLowerCase();
    if(!map[fgl]){var info=batchInfo[fgl]||{qty:1,unit:'Kg'};map[fgl]={name:fg2,qty:info.qty||1,unit:info.unit||'Kg',ingredients:[]};}
    var mat=trim(row[MAT]);if(!mat)continue;
 
    var pmEntry = pm[kw(mat) + '_subrm'] || pm[kw(mat)] || null;
    var rate=pmEntry&&pmEntry.price>0?pmEntry.price:(pn(row[RATE])>0?pn(row[RATE]):0);
    map[fgl].ingredients.push({material:mat,qty:pn(row[QTY]),unit:trim(row[UNIT]),rate:rate,lineTotal:pn(row[QTY])*rate});
  }
  return map;
}
 
function getRecipeMap(){
  try{return buildRecMap(rd('Recipes').raw,readAllRM());}catch(e){return{};}
}
 
function buildSubRmMap(){
  var map={};
  try{
    var sheet=SS().getSheetByName('Sub_RM_Recipes');if(!sheet)return{};
    var data=sheet.getDataRange().getValues();if(data.length<2)return{};
    var h=data[0].map(function(x){return trim(String(x)).toLowerCase();});
    var iN=h.indexOf('sub_rm_name'),iI=h.indexOf('ingredient_name'),iQ=h.indexOf('qty_required'),iU=h.indexOf('unit');
    var iOQ=h.indexOf('batch_output_qty');if(iOQ<0)iOQ=h.indexOf('batch output qty');
    var iOU=h.indexOf('batch_output_unit');if(iOU<0)iOU=h.indexOf('batch output unit');
    if(iN<0)iN=0;if(iI<0)iI=1;if(iQ<0)iQ=2;if(iU<0)iU=3;if(iOQ<0)iOQ=4;if(iOU<0)iOU=5;
    for(var i=1;i<data.length;i++){
      var row=data[i],name=row[iN]?trim(String(row[iN])):'';if(!name)continue;
      if(!map[name]){map[name]={name:name,outputQty:pn(row[iOQ])||1,outputUnit:row[iOU]?trim(String(row[iOU])):'Kg',ingredients:[]};}
      if(row[iI])map[name].ingredients.push({material:trim(String(row[iI])),qty:pn(row[iQ]),unit:row[iU]?trim(String(row[iU])):''});
    }
  }catch(e){Logger.log('buildSubRmMap:'+e.message);}
  return map;
}
 
function saveDualRecipe(fgName,ings,batchQty,batchUnit){
  try{
    var cleanName=trim(fgName),bQty=pn(batchQty)||1,bUnit=trim(batchUnit)||'Kg';
    if(!cleanName)throw new Error('Product name required');
    if(!ings||!ings.length)throw new Error('No ingredients provided');
    var pm=_buildPriceMap(),totalBatchCost=0;
    ings.forEach(function(g){
      var entry = pm[kw(trim(g.material)) + '_subrm'] || pm[kw(trim(g.material))] || {price:0};
      var r = pn(g.rate) > 0 ? pn(g.rate) : entry.price;
      totalBatchCost+=pn(g.qty)*r;
    });
    var normalizedCpu=bQty>0?totalBatchCost/bQty:totalBatchCost;
    var srmSh=ensureSh('Sub_RM_Recipes',['Sub_RM_Name','Ingredient_Name','Qty_Required','Unit','Batch_Output_Qty','Batch_Output_Unit']);
    var srmData=srmSh.getDataRange().getValues();
    for(var i=srmData.length-1;i>=1;i--){if(kw(trim(String(srmData[i][0]||'')))===kw(cleanName))srmSh.deleteRow(i+1);}
    ings.forEach(function(g,idx){var mat=trim(g.material),qty=pn(g.qty),unit=trim(g.unit)||'Kg';if(!mat)return;srmSh.appendRow([cleanName,mat,qty,unit,idx===0?bQty:'',idx===0?bUnit:'']);});
    saveRecipe(cleanName,ings,bQty,bUnit,true);
    ensureSh('FG_List',['FG_Name','Category','GST_Pct','Notes']);
    var fgD=rd('FG_List'),fgNi=fgD.h.indexOf('FG_Name');if(fgNi<0)fgNi=0;
    var fgExists=fgD.raw.slice(fgD.ds).some(function(r){return kw(trim(String(r[fgNi]||'')))===kw(cleanName);});
    if(!fgExists)fgD.s.appendRow([cleanName,'Sub-RM',5,'Auto-generated Sub-RM']);
    var rmSh=ensureRMSheet();
    var existRow=_findRowInRMWithCat(cleanName, 'Sub-RM');
    if(existRow>0){rmSh.getRange(existRow,RM_COL_PRICE,1,1).setValue(normalizedCpu);}
    else{rmSh.appendRow([cleanName,'','Sub-RM',bUnit,normalizedCpu]);}
    SpreadsheetApp.flush();
    return{success:true,message:'"'+cleanName+'" saved. Cost/'+bUnit+': ₹'+normalizedCpu.toFixed(2),cpu:normalizedCpu};
  }catch(e){throw new Error('Save Failed: '+e.message);}
}
 
function saveSubRM(subName,ings,replace,batchQty,batchUnit){
  subName=trim(subName);batchQty=pn(batchQty)||1;batchUnit=trim(batchUnit)||'Kg';
  if(!subName)throw new Error('Sub-RM Name required');
  if(!ings||!ings.length)throw new Error('No ingredients provided');
  var s=ensureSh('Sub_RM_Recipes',['Sub_RM_Name','Ingredient_Name','Qty_Required','Unit','Batch_Output_Qty','Batch_Output_Unit']);
  var d=rd('Sub_RM_Recipes'),ni=d.h.indexOf('Sub_RM_Name');if(ni<0)ni=0;
  if(replace){for(var i=d.raw.length-1;i>=d.ds;i--){if(kw(trim(String(d.raw[i][ni]||'')))===kw(subName))s.deleteRow(i+1);}}
  ings.forEach(function(g,idx){var m=trim(g.material),u=trim(g.unit)||'Kg',q=pn(g.qty);if(!m)return;s.appendRow([subName,m,q,u,idx===0?batchQty:'',idx===0?batchUnit:'']);});
  var pm=_buildPriceMap(),totalBatchCost=0;
  ings.forEach(function(g){
    var entry = pm[kw(trim(g.material)) + '_subrm'] || pm[kw(trim(g.material))] || {price:0};
    totalBatchCost += pn(g.qty) * entry.price;
  });
  var normalizedCpu=batchQty>0?totalBatchCost/batchQty:totalBatchCost;
  var rmSh=ensureRMSheet();
  var existRow=_findRowInRMWithCat(subName, 'Sub-RM');
  if(existRow>0){rmSh.getRange(existRow,RM_COL_PRICE,1,1).setValue(normalizedCpu);}
  else{
    rmSh.appendRow([subName,'','Sub-RM',batchUnit,normalizedCpu]);
    rmSh.getRange(rmSh.getLastRow(),1,1,5).setFontWeight('bold').setFontColor('#000000').setBackground(null);
  }
  SpreadsheetApp.flush();
  return{success:true,message:'"'+subName+'" Sub-RM saved. Cost/Unit: ₹'+normalizedCpu.toFixed(2),cpu:normalizedCpu};
}
 
function deleteSubRM(name){
  name=trim(name);
  var d=rd('Sub_RM_Recipes'),ni=d.h.indexOf('Sub_RM_Name');if(ni<0)ni=0;
  var n=0;
  for(var i=d.raw.length-1;i>=d.ds;i--){if(kw(trim(String(d.raw[i][ni]||'')))===kw(name)){d.s.deleteRow(i+1);n++;}}
  if(!n)throw new Error('"'+name+'" not found');
  try{var er=_findRowInRMWithCat(name, 'Sub-RM');if(er>0)ensureRMSheet().deleteRow(er);}catch(e){}
  SpreadsheetApp.flush();
  return{success:true,message:'"'+name+'" deleted.'};
}
 
function deleteRecipe(fgName){
  fgName=trim(fgName);var d=rd('Recipes'),fi=ci(d.h,['fg_name','fgname']);if(fi<0)fi=0;
  var n=0;for(var i=d.raw.length-1;i>=d.ds;i--){if(kw(trim(String(d.raw[i][fi]||'')))===kw(fgName)){d.s.deleteRow(i+1);n++;}}
  if(!n)throw new Error('"'+fgName+'" not found');
  SpreadsheetApp.flush();
  return{success:true,message:'Deleted.'};
}
 
function saveFG(fg){
  var name=trim(fg.name);if(!name)throw new Error('Name required');
  ensureSh('FG_List',['FG_Name','Category','GST_Pct','Notes']);
  var d=rd('FG_List'),ni=d.h.indexOf('FG_Name'),row=[name,trim(fg.category),pn(fg.gst),trim(fg.notes)];
  for(var i=d.ds;i<d.raw.length;i++){if(kw(trim(String(d.raw[i][ni]||'')))===kw(name)){d.s.getRange(i+1,1,1,4).setValues([row]);SpreadsheetApp.flush();return{success:true,message:'"'+name+'" updated.'};}}
  d.s.appendRow(row);SpreadsheetApp.flush();return{success:true,message:'"'+name+'" added.'};
}
 
function deleteFG(name){
  name=trim(name);var d=rd('FG_List'),ni=d.h.indexOf('FG_Name');
  for(var i=d.raw.length-1;i>=d.ds;i--){if(kw(trim(String(d.raw[i][ni]||'')))===kw(name)){d.s.deleteRow(i+1);SpreadsheetApp.flush();return{success:true,message:'Deleted.'};}}
  throw new Error('"'+name+'" not found');
}
 
// ════════════════════════════════════════════════════════════════════
// CALCULATE COST — COMPREHENSIVE ENGINE
// ════════════════════════════════════════════════════════════════════
function calculateCost(fgName,sellingPrice,ohOvr){
  sellingPrice=pn(sellingPrice);ohOvr=ohOvr||{};
  ensureRMSheet();
  var rmList=readAllRM(),ohRows=_readOHRows();
  var subMap=buildSubRmMap();
 
  var pm={};
  rmList.forEach(function(r){
    if(r.name){
      var e={price:r.price,unit:r.unit};
      pm[kw(r.name) + '_' + kw(r.category)] = e;
      if(!pm[kw(r.name)]) {
        pm[kw(r.name)] = e;
      }
    }
  });
 
  Object.keys(subMap).forEach(function(sn){
    var t=0,oq=pn(subMap[sn].outputQty)||1;
    subMap[sn].ingredients.forEach(function(g){
      var entry = pm[kw(g.material) + '_subrm'] || pm[kw(g.material)] || {price:0};
      t += pn(g.qty) * entry.price;
    });
    var cpu=t/oq;
    var masterSubEntry={price:cpu,unit:subMap[sn].outputUnit||'Kg'};
    pm[kw(sn) + '_subrm'] = masterSubEntry;
    pm[sn.toLowerCase().trim() + '_subrm'] = masterSubEntry;
  });
 
  var rMap=buildRecMap(rd('Recipes').raw,rmList);
 
  var fgKey=kw(fgName),rk=Object.keys(rMap).find(function(k){return kw(k)===fgKey;});
  if(!rk)throw new Error('No recipe for "'+fgName+'". Add recipe first.');
  var recipe=rMap[rk],fgMeta={gst:0,category:''};
  try{
    var fl=rd('FG_List').rows.find(function(r){return kw(trim(r['FG_Name']))===fgKey;});
    if(fl){fgMeta.gst=pn(fl['GST_Pct']);fgMeta.category=trim(fl['Category']);}
  }catch(e){}
 
  var lines=[],batchRMTotal=0,miss=[];
  recipe.ingredients.forEach(function(g){
    var entry = pm[kw(g.material) + '_subrm'] || pm[kw(g.material)] || null;
    var rate,unit;
    if(entry&&entry.price>0){
      rate=entry.price; unit=entry.unit||g.unit||'';
    }else if(entry){
      rate=pn(g.rate)>0?pn(g.rate):0; unit=entry.unit||g.unit||'';
    }else{
      rate=pn(g.rate)>0?pn(g.rate):0; unit=g.unit||'';
    }
    var line=pn(g.qty)*rate;
    if(rate<=0)miss.push(g.material);
    batchRMTotal+=line;
    lines.push({material:g.material,qty:pn(g.qty),unit:unit,rate:rate,line:line,
                missing:rate<=0,isSubRM:!!(subMap[g.material]||subMap[kw(g.material)])});
  });
 
  var bq=pn(recipe.qty)||1,cpRMPerUnit=batchRMTotal/bq;
  var ohL=[],ohTotalPerUnit=0;
  ohRows.forEach(function(o){
    var cat=trim(o.category),subCat=trim(o.subCategory||'');
    if(subCat&&kw(subCat)!==kw(fgMeta.category))return;
    var overrideKey=ckGS(cat)+(subCat?'_'+ckGS(subCat):'');
    var pct=(ohOvr[overrideKey]!==undefined&&ohOvr[overrideKey]!=='')?pn(ohOvr[overrideKey]):o.percentage;
    var basis=trim(o.basis)||'% of RM Cost';
    var baseValue=basis.toLowerCase().indexOf('selling')>-1?sellingPrice:cpRMPerUnit;
    var amt=baseValue*pct/100;ohTotalPerUnit+=amt;
    ohL.push({category:cat,subCategory:subCat,basis:basis,pct:pct,base:baseValue,amount:amt});
  });
 
  var subTotal=cpRMPerUnit+ohTotalPerUnit,gstAmt=subTotal*fgMeta.gst/100;
  var profitPU=sellingPrice>0?sellingPrice-subTotal:0,netPct=sellingPrice>0?(profitPU/sellingPrice)*100:0;
 
  return{
    fgName:fgName,fgCategory:fgMeta.category,
    batchQty:bq,batchUnit:recipe.unit||'Kg',
    lineItems:lines,missing:miss,batchRM:batchRMTotal,cpRMPerUnit:cpRMPerUnit,
    ohLines:ohL,ohTotal:ohTotalPerUnit,finCostPerUnit:subTotal,
    gstPct:fgMeta.gst,gstAmt:gstAmt,totalWithGst:subTotal+gstAmt,
    costPerUnit:subTotal,sellingPrice:sellingPrice,profitPerUnit:profitPU,
    netProfitPct:netPct,
    timestamp:new Date().toLocaleString('en-IN',{timeZone:'Asia/Kolkata'})
  };
}
 
function saveCalculation(r){
  var s=ensureSh('Cost_Log',['Timestamp','Product','Category','Batch Qty','Unit','Batch RM','Overhead','GST','Total Cost','Cost/Unit','Selling/Unit','Profit/Unit','Margin%','Net Profit%']);
  var mp=r.marginPct||r.netProfitPct||0;
  s.appendRow([r.timestamp,r.fgName,r.fgCategory||'',r.batchQty,r.batchUnit,
    +r.batchRM.toFixed(2),+r.ohTotal.toFixed(2),+r.gstAmt.toFixed(2),
    +r.totalWithGst.toFixed(2),+r.costPerUnit.toFixed(2),
    r.sellingPrice,+(r.profitPerUnit||0).toFixed(2),
    mp.toFixed(1)+'%',(r.netProfitPct||0).toFixed(1)+'%']);
  SpreadsheetApp.flush();
  return true;
}
 
function saveOverhead(cat,subCat,basis,pct){
  cat=trim(cat);subCat=trim(subCat);basis=trim(basis)||'% of RM Cost';pct=pn(pct);
  if(!cat)throw new Error('Category required');
  ensureSh('Overheads',['Category','Sub_Category','Calculation_Basis','Percentage']);
  var d=rd('Overheads'),ci2=d.h.indexOf('Category'),sci=d.h.indexOf('Sub_Category');if(ci2<0)ci2=0;
  for(var i=d.ds;i<d.raw.length;i++){
    var rc=kw(trim(String(d.raw[i][ci2]||''))),rs=kw(trim(String(sci>=0?d.raw[i][sci]:'')||''));
    if(rc===kw(cat)&&rs===kw(subCat)){d.s.getRange(i+1,1,1,4).setValues([[cat,subCat,basis,pct]]);SpreadsheetApp.flush();return{success:true,message:'Updated.',overheads:_readOHRows()};}
  }
  d.s.appendRow([cat,subCat,basis,pct]);SpreadsheetApp.flush();return{success:true,message:'Added.',overheads:_readOHRows()};
}
 
function updateOverhead(origCat,origSubCat,nc,nSubCat,basis,pct){
  origCat=trim(origCat);origSubCat=trim(origSubCat);nc=trim(nc);nSubCat=trim(nSubCat);
  basis=trim(basis)||'% of RM Cost';pct=pn(pct);if(!nc)throw new Error('Name required');
  var d=rd('Overheads'),ci2=d.h.indexOf('Category'),sci=d.h.indexOf('Sub_Category');if(ci2<0)ci2=0;
  for(var i=d.ds;i<d.raw.length;i++){
    var rc=kw(trim(String(d.raw[i][ci2]||''))),rs=kw(trim(String(sci>=0?d.raw[i][sci]:'')||''));
    if(rc===kw(origCat)&&rs===kw(origSubCat)){d.s.getRange(i+1,1,1,4).setValues([[nc,nSubCat,basis,pct]]);SpreadsheetApp.flush();return{success:true,message:'Updated.',overheads:_readOHRows()};}
  }
  d.s.appendRow([nc,nSubCat,basis,pct]);SpreadsheetApp.flush();return{success:true,message:'Saved.',overheads:_readOHRows()};
}
 
function deleteOverhead(cat,subCat){
  cat=trim(cat);subCat=trim(subCat);
  var d=rd('Overheads'),ci2=d.h.indexOf('Category'),sci=d.h.indexOf('Sub_Category');if(ci2<0)ci2=0;
  for(var i=d.raw.length-1;i>=d.ds;i--){
    var rc=kw(trim(String(d.raw[i][ci2]||''))),rs=kw(trim(String(sci>=0?d.raw[i][sci]:'')||''));
    if(rc===kw(cat)&&rs===kw(subCat)){d.s.deleteRow(i+1);return{success:true,message:'Deleted.',overheads:_readOHRows()};}
  }
  throw new Error('"'+cat+'" not found');
}
 
// ════════════════════════════════════════════════════════════════════
// RATE MANAGER FUNCTIONS
// ════════════════════════════════════════════════════════════════════
//
// ALL category sheets (Mithai Raw, Nasta Raw, Namkeen Raw, Savli Bakery)
// AND RM_Price share the SAME 5-column flat format:
//   A: Material_Name_Eng
//   B: Material_Name_Hindi
//   C: Category
//   D: Unit
//   E: Price_Per_Unit
// NO GST column in any of these sheets.
//
// ════════════════════════════════════════════════════════════════════
 
/**
 * Reads a category sheet or RM_Price for the Rate Manager UI (NewIndex.html).
 *
 * Category sheets (Nasta Raw, Namkeen Raw, Savli Bakery, Mithai Raw) use DUAL-BLOCK layout:
 *   Block 1: col A=Hindi, col B=Eng, col C=Rate
 *   Block 2: col G=Eng,   col H=Rate  (no Hindi)
 *   Column positions come from CAT_SHEET_CONFIG.sets (1-based)
 *
 * RM_Price uses FLAT 5-column layout:
 *   A=Material_Name_Eng, B=Material_Name_Hindi, C=Category, D=Unit, E=Price_Per_Unit
 */
function getSheetData(sheetName) {
  try {
    sheetName = trim(sheetName);
 
    var ss = SS();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      try { ss = SpreadsheetApp.openById(SPREADSHEET_ID); sheet = ss.getSheetByName(sheetName); } catch(fe){}
    }
    if (!sheet) {
      Logger.log('getSheetData: sheet not found: "' + sheetName + '"');
      return [];
    }
 
    var last = sheet.getLastRow();
    if (last < 2) return [];
 
    var processedData = [];
    var seen = {};
 
    // ── RM_Price: always flat 5-column ──
    // A=Material_Name_Eng | B=Material_Name_Hindi | C=Category | D=Unit | E=Price_Per_Unit
    if (sheetName === RM_SHEET) {
      var vals = sheet.getRange(2, 1, last - 1, 5).getValues();
      for (var i = 0; i < vals.length; i++) {
        var row = vals[i];
        var eng = trim(String(row[0] || ''));
        if (!eng || /[^\x00-\x7F]/.test(eng)) continue;
        var k = eng.toLowerCase();
        if (seen[k]) continue;
        seen[k] = true;
        var cat = trim(String(row[2] || ''));
        processedData.push({
          nameEnglish : eng,
          nameHindi   : trim(String(row[1] || '')),
          category    : cat,
          unit        : trim(String(row[3] || '')) || 'Kg',
          rate        : parseFloat(row[4]) || 0,
          gst         : GST[cat] !== undefined ? GST[cat] : 0,
          rowIdx      : i + 2,
          colStart    : 5
        });
      }
      return processedData;
    }
 
    // ── Category sheets: use CAT_SHEET_CONFIG to determine format ──
    var cfg = CAT_SHEET_CONFIG[sheetName];
    if (!cfg) {
      var cfgKey = Object.keys(CAT_SHEET_CONFIG).filter(function(k) {
        return trim(k) === sheetName;
      })[0];
      if (cfgKey) cfg = CAT_SHEET_CONFIG[cfgKey];
    }
    if (!cfg) {
      Logger.log('getSheetData: no config for sheet "' + sheetName + '"');
      return [];
    }
 
    var categoryGst = GST[cfg.cat] !== undefined ? GST[cfg.cat] : 0;
 
    if (cfg.flat) {
      // ── FLAT 5-column format (e.g. Mithai Raw) ──
      // A=Eng, B=Hindi, C=Category, D=Unit, E=Price_Per_Unit
      var vals = sheet.getRange(2, 1, last - 1, 5).getValues();
      for (var i = 0; i < vals.length; i++) {
        var row = vals[i];
        var eng = trim(String(row[0] || ''));
        if (!eng || /[^\x00-\x7F]/.test(eng)) continue;
        var k = eng.toLowerCase();
        if (seen[k]) continue;
        seen[k] = true;
        processedData.push({
          nameEnglish : eng,
          nameHindi   : trim(String(row[1] || '')),
          category    : cfg.cat,
          unit        : trim(String(row[3] || '')) || 'Kg',
          rate        : parseFloat(row[4]) || 0,
          gst         : categoryGst,
          rowIdx      : i + 2,
          colStart    : 5
        });
      }
    } else {
      // ── DUAL-BLOCK format (e.g. Nasta Raw, Namkeen Raw, Savli Bakery) ──
      // Block 1: A=Hindi(1), B=Eng(2), C=Rate(3)
      // Block 2: G=Eng(7),   H=Rate(8)  ← no Hindi col in Block 2
      var maxCol = 9;
      var vals = sheet.getRange(1, 1, last, maxCol).getValues();
 
      // Build Hindi lookup from Block 1 sets (where hinCol is available)
      // So Block 2 items can also get their Hindi name
      var hindiLookup = {};
      cfg.sets.forEach(function(set) {
        if (set.hinCol == null) return;
        var ec = set.engCol - 1;
        var hc = set.hinCol - 1;
        for (var i = 1; i < vals.length; i++) {
          var eng = trim(String(vals[i][ec] || ''));
          var hin = trim(String(vals[i][hc] || ''));
          if (eng && !/[^\x00-\x7F]/.test(eng)) {
            hindiLookup[eng.toLowerCase()] = hin || hindiLookup[eng.toLowerCase()] || '';
          }
        }
      });
 
      cfg.sets.forEach(function(set) {
        var ec = set.engCol  - 1;
        var hc = set.hinCol != null ? set.hinCol - 1 : -1;
        var rc = set.rateCol - 1;
        for (var i = 1; i < vals.length; i++) {
          var row = vals[i];
          var eng = trim(String(row[ec] || ''));
          if (!eng || /[^\x00-\x7F]/.test(eng)) continue;
          var k = eng.toLowerCase();
          if (seen[k]) continue;
          seen[k] = true;
          // Hindi: from same row (Block 1), or cross-lookup from Block 1 hindiLookup (Block 2)
          var hin = hc >= 0 ? trim(String(row[hc] || '')) : (hindiLookup[k] || '');
          processedData.push({
            nameEnglish : eng,
            nameHindi   : hin,
            category    : cfg.cat,
            unit        : 'Kg',
            rate        : parseFloat(row[rc]) || 0,
            gst         : categoryGst,
            rowIdx      : i + 1,
            colStart    : set.rateCol
          });
        }
      });
 
      // Fallback: fill still-missing Hindi names from RM_Price
      var rmAll = _readRMAll();
      var rmHindiMap = {};
      rmAll.forEach(function(r) { if (r.name) rmHindiMap[r.name.toLowerCase()] = r.nameHi || ''; });
      processedData.forEach(function(item) {
        if (!item.nameHindi && rmHindiMap[item.nameEnglish.toLowerCase()]) {
          item.nameHindi = rmHindiMap[item.nameEnglish.toLowerCase()];
        }
      });
    }
 
    return processedData;
  } catch (e) {
    Logger.log('Error in getSheetData: ' + e.message);
    return [];
  }
}
function updateSheetData(sheetName, updatedItems, isGlobal) {
  try {
    var ss = SS();
    if (!ss.getSheetByName(DEPT_SHEETS[0])) {
      try { ss = SpreadsheetApp.openById(SPREADSHEET_ID); } catch(fe){}
    }
    var timestamp = new Date();
 
    // ── History sheet ──
    var hist = ss.getSheetByName('Update History') || ss.insertSheet('Update History');
    if (hist.getLastRow() === 0) {
      hist.appendRow(['Timestamp','Type','Source','Synced To','English Name','Hindi Name','New Rate','Category']);
    }
 
    // ── Reverse lookup: sheet name → Category string ──
    var SHEET_TO_CAT = {};
    Object.keys(CAT_SHEET_CONFIG).forEach(function(sn){ SHEET_TO_CAT[sn] = CAT_SHEET_CONFIG[sn].cat; });
 
    // ── Load full RM_Price into memory once ──
    // rmMap: kw(engName)+'_'+kw(category) → { idx: 0-based index, rowNum: 1-based sheet row }
    var rmSh   = ensureRMSheet();
    var rmLast = rmSh.getLastRow();
    var rmData = rmLast >= RM_DATA_START
      ? rmSh.getRange(RM_DATA_START, 1, rmLast - RM_DATA_START + 1, 5).getValues()
      : [];
    var rmMap = {};
    rmData.forEach(function(r, idx) {
      var n   = trim(String(r[0] || ''));
      var cat = trim(String(r[2] || ''));
      if (n) {
        // Primary key: name + category (strict match)
        rmMap[kw(n) + '_' + kw(cat)] = { idx: idx, rowNum: RM_DATA_START + idx };
        // Fallback key: name only (used if category not known)
        if (!rmMap['_name_' + kw(n)]) rmMap['_name_' + kw(n)] = { idx: idx, rowNum: RM_DATA_START + idx };
      }
    });
 
    // ── Load Recipes into memory once ──
    // recipeMatMap: kw(matName)+'_'+kw(category) → [row indices in recData]
    var recSh   = SS().getSheetByName('Recipes');
    var recData = [], recRange = null;
    if (recSh && recSh.getLastRow() >= 2) {
      recRange = recSh.getRange(2, 1, recSh.getLastRow() - 1, 8);
      recData  = recRange.getValues();
    }
    // Build FG→category map from FG_List for recipe category matching
    var fgCatMap = {};
    try {
      var fgD = rd('FG_List');
      var fgNi = fgD.h.indexOf('FG_Name');
      var fgCi = fgD.h.indexOf('Category');
      if (fgNi >= 0 && fgCi >= 0) {
        fgD.rows.forEach(function(r) {
          var fgName = trim(String(r['FG_Name'] || ''));
          var fgCat  = trim(String(r['Category'] || ''));
          if (fgName) fgCatMap[fgName.toLowerCase()] = fgCat;
        });
      }
    } catch(fe) { Logger.log('fgCatMap build error: ' + fe.message); }
 
    var newRMRows     = [];
    var historyRows   = [];
    var recipeUpdated = 0;
 
    updatedItems.forEach(function(item) {
      var engName  = trim(String(item.nameEnglish || ''));
      var hinName  = trim(String(item.nameHindi   || ''));
      var newRate  = pn(item.rate);
      var unit     = trim(String(item.unit || '')) || 'Kg';
      var needle   = engName.toLowerCase();
      var syncedNames = [];
 
      if (!engName) return;
 
      // Derive category from source sheet
      var itemCategory = item.category || SHEET_TO_CAT[sheetName] || '';
 
      // ════════════════════════════════════════
      // STEP 1 — Update rate in category sheet(s)
      // Match: name AND category must match
      // ════════════════════════════════════════
      if (isGlobal) {
        DEPT_SHEETS.forEach(function(sName) {
          var tSh  = ss.getSheetByName(sName);
          if (!tSh) return;
          var tLast = tSh.getLastRow();
          if (tLast < 2) return;
          var tCfg = CAT_SHEET_CONFIG[sName];
          if (!tCfg) {
            var tk = Object.keys(CAT_SHEET_CONFIG).filter(function(k){ return trim(k) === trim(sName); })[0];
            if (tk) tCfg = CAT_SHEET_CONFIG[tk];
          }
          if (!tCfg) return;
 
          // Category of this sheet must match item's category
          if (kw(tCfg.cat) !== kw(itemCategory)) return;
 
          var modified = false;
          if (tCfg.flat) {
            // Flat: col A=Eng, col C=Category, col E=Price — match name + category
            var tData = tSh.getRange(2, 1, tLast - 1, 5).getValues();
            for (var r = 0; r < tData.length; r++) {
              var rowEng = trim(String(tData[r][0] || '')).toLowerCase();
              var rowCat = kw(trim(String(tData[r][2] || '')));
              if (rowEng === needle && rowCat === kw(itemCategory)) {
                tSh.getRange(r + 2, 5).setValue(newRate);
                modified = true;
              }
            }
          } else {
            // Dual-block: col B=Eng, col C=Rate (Block1) | col G=Eng, col H=Rate (Block2)
            // Category matched at sheet level already — match by eng name only within sheet
            var tData = tSh.getRange(1, 1, tLast, 9).getValues();
            tCfg.sets.forEach(function(set) {
              var ec = set.engCol - 1;
              var rc = set.rateCol;     // 1-based
              for (var r = 1; r < tData.length; r++) {
                if (trim(String(tData[r][ec] || '')).toLowerCase() === needle) {
                  tSh.getRange(r + 1, rc).setValue(newRate);
                  modified = true;
                }
              }
            });
          }
          if (modified) syncedNames.push(sName);
        });
      } else {
        // Local update: verify name + category match at rowIdx before writing
        var tSh = ss.getSheetByName(sheetName);
        if (tSh && item.rowIdx && item.colStart) {
          var tCfg = CAT_SHEET_CONFIG[sheetName];
          if (!tCfg) {
            var tk2 = Object.keys(CAT_SHEET_CONFIG).filter(function(k){ return trim(k) === trim(sheetName); })[0];
            if (tk2) tCfg = CAT_SHEET_CONFIG[tk2];
          }
          // Verify the row still has the expected eng name (prevents stale rowIdx writes)
          var verifyOk = false;
          if (tCfg) {
            var verifyRow = tSh.getRange(item.rowIdx, 1, 1, 9).getValues()[0];
            if (tCfg.flat) {
              // Flat: col A=Eng, col C=Category
              verifyOk = trim(String(verifyRow[0] || '')).toLowerCase() === needle
                      && kw(trim(String(verifyRow[2] || ''))) === kw(itemCategory);
            } else {
              // Dual-block: check engCol
              var anyMatch = tCfg.sets.some(function(set) {
                return trim(String(verifyRow[set.engCol - 1] || '')).toLowerCase() === needle;
              });
              verifyOk = anyMatch;
            }
          }
          if (verifyOk) {
            tSh.getRange(item.rowIdx, item.colStart).setValue(newRate);
            syncedNames.push(sheetName);
          } else {
            // rowIdx mismatch — scan sheet to find correct row
            var tLast2 = tSh.getLastRow();
            if (tLast2 >= 2 && tCfg) {
              if (tCfg.flat) {
                var tData2 = tSh.getRange(2, 1, tLast2 - 1, 5).getValues();
                for (var r2 = 0; r2 < tData2.length; r2++) {
                  if (trim(String(tData2[r2][0] || '')).toLowerCase() === needle
                   && kw(trim(String(tData2[r2][2] || ''))) === kw(itemCategory)) {
                    tSh.getRange(r2 + 2, 5).setValue(newRate);
                    syncedNames.push(sheetName);
                    break;
                  }
                }
              } else {
                var tData2 = tSh.getRange(1, 1, tLast2, 9).getValues();
                tCfg.sets.forEach(function(set) {
                  for (var r2 = 1; r2 < tData2.length; r2++) {
                    if (trim(String(tData2[r2][set.engCol - 1] || '')).toLowerCase() === needle) {
                      tSh.getRange(r2 + 1, set.rateCol).setValue(newRate);
                      if (syncedNames.indexOf(sheetName) < 0) syncedNames.push(sheetName);
                    }
                  }
                });
              }
            }
          }
        }
      }
 
      // ════════════════════════════════════════
      // STEP 2 — RM_Price sync
      // Match: name + category (strict); fallback name-only if category not found
      // ════════════════════════════════════════
      var rmKey      = kw(engName) + '_' + kw(itemCategory);
      var rmEntry    = rmMap[rmKey] || rmMap['_name_' + kw(engName)];
 
      if (rmEntry !== undefined) {
        // Item exists in RM_Price — update Price_Per_Unit (col E) only if category matches
        var existingCat = kw(trim(String(rmData[rmEntry.idx][2] || '')));
        if (existingCat === kw(itemCategory) || itemCategory === '') {
          rmSh.getRange(rmEntry.rowNum, RM_COL_PRICE, 1, 1).setValue(newRate);
        }
      } else {
        // New item — queue for append with correct category
        var srcSheet = isGlobal ? (syncedNames[0] || sheetName) : sheetName;
        var category = itemCategory || SHEET_TO_CAT[srcSheet] || 'Other';
        newRMRows.push([engName, hinName, category, unit, newRate]);
        // Register in rmMap so duplicate in same batch is skipped
        rmMap[kw(engName) + '_' + kw(category)] = { idx: rmData.length + newRMRows.length - 1, rowNum: -1 };
        rmMap['_name_' + kw(engName)]            = rmMap[kw(engName) + '_' + kw(category)];
      }
 
      // ════════════════════════════════════════
      // STEP 3 — Recipes sync
      // Match: Material_Name (kw) — update Unit_Rate + Line_Total
      // Category check already enforced in STEP 1 (source sheet) & STEP 2 (RM_Price)
      // ════════════════════════════════════════
      if (recData.length > 0) {
        for (var ri = 0; ri < recData.length; ri++) {
          var matName = trim(String(recData[ri][1] || ''));  // col B = Material_Name
          if (!matName) continue;
          if (kw(matName) !== kw(engName)) continue;        // name must match (keyword normalised)
          var qty = pn(recData[ri][2]);                     // col C = Qty_Required
          recData[ri][4] = newRate;                         // col E = Unit_Rate
          recData[ri][5] = qty * newRate;                   // col F = Line_Total = Qty × Rate
          recipeUpdated++;
        }
      }
 
      historyRows.push([timestamp, isGlobal ? 'Global' : 'Local', sheetName,
                        syncedNames.join(', '), engName, hinName, newRate, itemCategory]);
    });
 
    // ── Batch write Recipes changes ──
    if (recipeUpdated > 0 && recRange) {
      recRange.setValues(recData);
    }
 
    // ── Batch append new RM_Price rows ──
    if (newRMRows.length > 0) {
      var rmAppendRow = rmSh.getLastRow() + 1;
      var rmNewRange  = rmSh.getRange(rmAppendRow, 1, newRMRows.length, 5);
      rmNewRange.setValues(newRMRows);
      rmNewRange.setFontWeight('bold').setFontColor('#000000').setBackground(null);
    }
 
    // ── History log ──
    if (historyRows.length > 0) {
      hist.getRange(hist.getLastRow() + 1, 1, historyRows.length, 8).setValues(historyRows);
    }
 
    SpreadsheetApp.flush();
 
    var addedMsg  = newRMRows.length  > 0 ? ' | ' + newRMRows.length  + ' new item(s) added to RM_Price' : '';
    var recipeMsg = recipeUpdated     > 0 ? ' | ' + recipeUpdated      + ' recipe line(s) updated'        : '';
    return '✅ Updated ' + updatedItems.length + ' item(s).' + addedMsg + recipeMsg;
  } catch (e) {
    return '❌ Error: ' + e.message;
  }
}
 
// ════════════════════════════════════════════════════════════════════
// MAIN ENTRY POINT — UNIFIED doGet ROUTER
// ════════════════════════════════════════════════════════════════════
// Routes:
//   ?role=staff  → Staff.html   (MithaiCalc Staff View)
//   ?role=rates  → NewIndex.html (Rate Manager)
//   (default)    → Index.html   (MithaiCalc Admin)
// ════════════════════════════════════════════════════════════════════
function doGet(e) {
  ensureRMSheet();
  var role = (e && e.parameter && e.parameter.role) ? e.parameter.role.toLowerCase() : '';
 
  if (role === 'staff') {
    return HtmlService.createHtmlOutputFromFile('Staff')
      .setTitle('MithaiCalc — Staff')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
 
  if (role === 'rates') {
    return HtmlService.createTemplateFromFile('NewIndex')
      .evaluate()
      .setTitle('Raghuvir Sweets - Rate Manager')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
 
  // Default → MithaiCalc Admin
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('MithaiCalc v13')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
