// MithaiCalc v7.8 — Overheads: Category|Sub_Category|Calculation_Basis|Percentage
function doGet(e){
  var role=(e&&e.parameter&&e.parameter.role)||'';
  if(role.toLowerCase()==='staff'){
    return HtmlService.createHtmlOutputFromFile('Staff').setTitle('MithaiCalc — Staff').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Cost Calculator').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

var GST={'Namkeen':5,'Methai':5,'Bakery':5};
var RM_CATS=['Namkeen','Methai','Bakery'];
var FG_CATS=['Methai','Bakery','Namkeen'];
var OH_SUB_CATS=['Methai','Bakery','Namkeen'];

// ── UTILITY FUNCTIONS ──
function pn(v){if(v===null||v===undefined||v==='')return 0;if(typeof v==='number')return isNaN(v)?0:v;var n=parseFloat(String(v).replace(/,/g,'').replace(/%/g,'').trim());return isNaN(n)?0:n;}
function trim(v){return String(v||'').trim();}
function isDec(s){s=trim(s);return!s||/^[\u{1F300}-\u{1FFFF}\u{2600}-\u{27BF}⚠ℹ🍬📋📊🧾]/u.test(s);}
function kw(s){return trim(s).toLowerCase().replace(/[^a-z0-9]/g,'');}
function SS(){return SpreadsheetApp.getActiveSpreadsheet();}
function sh(n){var s=SS().getSheetByName(n);if(!s)throw new Error('Sheet "'+n+'" not found');return s;}
function ensureSh(n,h){
  var ss=SS(),s=ss.getSheetByName(n);
  if(!s){
    s=ss.insertSheet(n);
    s.appendRow(h);
    s.getRange(1,1,1,h.length).setFontWeight('bold').setBackground('#3B1F0E').setFontColor('#FFF');
    s.setFrozenRows(1);
  }
  return s;
}

function rd(name){
  var s=sh(name),data=s.getDataRange().getValues();
  if(data.length<2)return{h:[],rows:[],ds:1,s:s,raw:data,hr:0};
  var hr=0;
  var KEYS=['fgname','materialname','ingredientname','category','unit','price','percentage','qty','product','overhead','cost','gst','subrmname','finalprod','subcategory'];
  for(var i=0;i<Math.min(5,data.length);i++){
    var row=data[i],score=0;
    row.forEach(function(c){var k=kw(c);KEYS.forEach(function(x){if(k.indexOf(x)>-1||x.indexOf(k)>-1)score++;});});
    if(score>=2){hr=i;break;}
  }
  var h=data[hr].map(function(x){return trim(x);});
  var rows=data.slice(hr+1).filter(function(r){var f=trim(r[0]);return f&&!isDec(f);}).map(function(r){var o={};h.forEach(function(k,i){o[k]=r[i]!==undefined?r[i]:'';});return o;});
  return{h:h,rows:rows,ds:hr+1,s:s,raw:data,hr:hr};
}

function ci(headers,keywords){
  var ks=headers.map(kw);
  for(var j=0;j<keywords.length;j++){
    for(var i=0;i<ks.length;i++){if(ks[i]===keywords[j]||ks[i].indexOf(keywords[j])>-1||keywords[j].indexOf(ks[i])>-1)return i;}
  }
  return -1;
}

// ── INTERNAL OVERHEAD HELPER ──
// Sheet columns: Category | Sub_Category | Calculation_Basis | Percentage
function _readOHRows(){
  try{
    return rd('Overheads').rows
      .filter(function(r){var c=trim(r['Category']);return c&&!c.toUpperCase().includes('TOTAL');})
      .map(function(r){return{
        category   : trim(r['Category']          || ''),
        subCategory: trim(r['Sub_Category']      || ''),
        basis      : trim(r['Calculation_Basis'] || '% of RM Cost'),
        percentage : pn(r['Percentage']          || 0)
      };});
  }catch(e){return[];}
}

// ── KEY SANITISER (mirrors frontend ck()) ──
function ckGS(c){return String(c||'').toLowerCase().replace(/[^a-z0-9]/g,'_').replace(/_+/g,'_').replace(/^_|_$/g,'');}

// ══════════════════════════════════════════════
// ── SAVE RECIPE (only saves to Recipes sheet) ──
// Called when user clicks "💾 Save Recipe"
// ══════════════════════════════════════════════
function saveRecipe(fgName, ings, batchQty, batchUnit, replace){
  fgName   = trim(fgName);
  batchQty = pn(batchQty) || 1;
  batchUnit= trim(batchUnit) || 'Kg';

  if(!fgName)       throw new Error('Product name is required');
  if(!ings || !ings.length) throw new Error('No ingredients provided');

  // Ensure Recipes sheet exists with correct headers
  ensureSh('Recipes',[
    'FG_Name','Material_Name','Qty_Required','Unit','Unit_Rate',
    'Final_Product_Qty','Final_Product_Unit'
  ]);

  var d  = rd('Recipes');
  var fi = ci(d.h, ['fgname','productname','finishedgood','product','fg']);
  if(fi < 0) fi = 0;

  // If replace = true, delete existing rows for this FG first
  if(replace){
    for(var i = d.raw.length - 1; i >= d.ds; i--){
      if(kw(trim(String(d.raw[i][fi] || ''))) === kw(fgName)){
        d.s.deleteRow(i + 1);
      }
    }
  }

  // Append new ingredient rows
  ings.forEach(function(g, idx){
    var mat  = trim(g.material);
    var qty  = pn(g.qty);
    var unit = trim(g.unit) || 'Kg';
    var rate = pn(g.rate)   || 0;
    if(!mat) return;
    // Store batch qty/unit only on first row to keep sheet clean
    d.s.appendRow([
      fgName, mat, qty, unit, rate,
      idx === 0 ? batchQty  : '',
      idx === 0 ? batchUnit : ''
    ]);
  });

  return { success: true, message: '"' + fgName + '" recipe saved successfully.' };
}

// ══════════════════════════════════════════════════════════════════════════════
// ── SAVE SUB-RAW MATERIAL RECIPE ──
// Called when user clicks "💾 Save Sub Raw Material Recipe"
// Saves to: Sub_RM_Recipes + Recipes + FG_List + RM_Price (normalized cost)
// ══════════════════════════════════════════════════════════════════════════════
function saveDualRecipe(fgName, ings, batchQty, batchUnit){
  try{
    var cleanName = trim(fgName);
    var bQty      = pn(batchQty) || 1;
    var bUnit     = trim(batchUnit) || 'Kg';

    if(!cleanName)       throw new Error('Product name is required');
    if(!ings || !ings.length) throw new Error('No ingredients provided');

    // ── STEP 1: Calculate Normalized Cost Per Unit ──
    var rmData   = sh('RM_Price').getDataRange().getValues();
    var priceMap = {};
    rmData.slice(1).forEach(function(r){
      var nm = trim(String(r[0] || ''));
      if(nm) priceMap[kw(nm)] = pn(r[4]);
    });

    var totalBatchCost = 0;
    ings.forEach(function(g){
      var rate = pn(g.rate) > 0 ? pn(g.rate) : (priceMap[kw(trim(g.material))] || 0);
      totalBatchCost += pn(g.qty) * rate;
    });
    var normalizedCpu = bQty > 0 ? totalBatchCost / bQty : totalBatchCost;

    // ── STEP 2: Save to Sub_RM_Recipes ──
    ensureSh('Sub_RM_Recipes',[
      'Sub_RM_Name','Ingredient_Name','Qty_Required','Unit','Batch_Output_Qty','Batch_Output_Unit'
    ]);
    var srmData = sh('Sub_RM_Recipes').getDataRange().getValues();
    var srmSh   = sh('Sub_RM_Recipes');

    // Delete existing rows for this Sub-RM
    for(var i = srmData.length - 1; i >= 1; i--){
      if(kw(trim(String(srmData[i][0] || ''))) === kw(cleanName)){
        srmSh.deleteRow(i + 1);
      }
    }
    // Append new rows
    ings.forEach(function(g, idx){
      var mat  = trim(g.material);
      var qty  = pn(g.qty);
      var unit = trim(g.unit) || 'Kg';
      if(!mat) return;
      srmSh.appendRow([
        cleanName, mat, qty, unit,
        idx === 0 ? bQty  : '',
        idx === 0 ? bUnit : ''
      ]);
    });

    // ── STEP 3: Save to Recipes sheet ──
    saveRecipe(cleanName, ings, bQty, bUnit, true);

    // ── STEP 4: Add to FG_List if not already present ──
    ensureSh('FG_List',['FG_Name','Category','GST_Pct','Notes']);
    var fgD  = rd('FG_List');
    var fgNi = fgD.h.indexOf('FG_Name');
    if(fgNi < 0) fgNi = 0;
    var fgExists = false;
    for(var j = fgD.ds; j < fgD.raw.length; j++){
      if(kw(trim(String(fgD.raw[j][fgNi] || ''))) === kw(cleanName)){
        fgExists = true; break;
      }
    }
    if(!fgExists){
      fgD.s.appendRow([cleanName, 'Sub-RM', 5, 'Auto-generated Sub-RM']);
    }

    // ── STEP 5: Update RM_Price with Normalized Cost ──
    var rmSh    = sh('RM_Price');
    var rd2     = rd('RM_Price');
    var ni2     = ci(rd2.h, ['materialnameeng','materialname']);
    if(ni2 < 0) ni2 = 0;
    var found   = false;
    for(var k = rd2.ds; k < rd2.raw.length; k++){
      if(kw(trim(String(rd2.raw[k][ni2] || ''))) === kw(cleanName)){
        rmSh.getRange(k + 1, 5).setValue(normalizedCpu);
        rmSh.getRange(k + 1, 4).setValue(bUnit);
        found = true; break;
      }
    }
    if(!found){
      rmSh.appendRow([cleanName, '', 'Sub-RM', bUnit, normalizedCpu]);
    }

    return {
      success : true,
      message : '"' + cleanName + '" saved to Sub-RM, Recipe, FG List & RM Price. Cost/Unit: ₹' + normalizedCpu.toFixed(2)
    };

  }catch(e){
    throw new Error('Save Failed: ' + e.message);
  }
}

// ══════════════════════════════════════
// ── GET ALL DATA ──
// ══════════════════════════════════════
function getData(){
  var fg=[],rm=[],rec={},subRm={},oh=[],recRaw=[];
  try{
    fg=rd('FG_List').rows.map(function(r){
      return{name:trim(r['FG_Name']),category:trim(r['Category']),gst:pn(r['GST_Pct']),notes:trim(r['Notes'])};
    }).filter(function(r){return r.name;});
  }catch(e){Logger.log('FG:'+e);}

  try{
    rm=rd('RM_Price').rows.map(function(r){
      return{
        name    : trim(r['Material_Name_Eng']||r['Material_Name']||''),
        nameHi  : trim(r['Material_Name_Hindi']||''),
        category: trim(r['Category']||''),
        unit    : trim(r['Unit']||''),
        price   : pn(r['Price_Per_Unit']||r['Price']||0)
      };
    }).filter(function(r){return r.name;});
  }catch(e){Logger.log('RM:'+e);}

  try{recRaw=rd('Recipes').raw;rec=buildRecMap(recRaw,rm);}catch(e){Logger.log('REC:'+e);}
  try{subRm=buildSubRmMap();}catch(e){Logger.log('SubRM:'+e);}
  oh=_readOHRows();

  return{
    fg:fg, rm:rm, recipes:rec, subRm:subRm, oh:oh,
    ohPct:oh.reduce(function(s,r){return s+r.percentage;},0),
    gst:GST, rmCats:RM_CATS, fgCats:FG_CATS, ohSubCats:OH_SUB_CATS,
    recCount:recRaw.length, rmCount:rm.length
  };
}

// ══════════════════════════════════════
// ── BUILD RECIPE MAP ──
// ══════════════════════════════════════
function buildRecMap(raw, rmList){
  var map={}, pm={};
  if(rmList&&rmList.length){
    rmList.forEach(function(r){if(r.name)pm[r.name.toLowerCase()]={price:r.price,unit:r.unit};});
  }
  if(!raw||raw.length<2)return map;
  var hr=0;
  for(var i=0;i<Math.min(5,raw.length);i++){
    var s=0;
    raw[i].forEach(function(c){var k=kw(c);if(['fgname','material','ingredient','qty','unit'].some(function(x){return k.indexOf(x)>-1;}))s++;});
    if(s>=2){hr=i;break;}
  }
  var h=raw[hr];
  var colIdx={};
  h.forEach(function(cell,idx){colIdx[kw(cell)]=idx;});
  function col(variants){for(var v=0;v<variants.length;v++){if(colIdx[variants[v]]!==undefined)return colIdx[variants[v]];}return -1;}
  var FG  =col(['fgname','productname','finishedgood','product','fg']);
  var MAT =col(['materialname','material','ingredientname','ingredient','rawmaterial']);
  var QTY =col(['qtyrequired','qty','quantity']);
  var UNIT=col(['unit','uom','measure']);
  var RATE=col(['unitrate','rate','priceperunit','unitprice','price']);
  var BQTY=col(['finalproductqty','finalprodqty','outputqty','yieldqty','piecesperbatch','pieces','batchqty']);
  var BUNT=col(['finalproductunit','finalunit','outputunit','yieldunit','batchunit']);
  if(FG<0)FG=0;if(MAT<0)MAT=1;if(QTY<0)QTY=2;if(UNIT<0)UNIT=3;
  if(RATE<0)RATE=4;if(BQTY<0)BQTY=5;if(BUNT<0)BUNT=6;

  var batchInfo={};
  for(var r=hr+1;r<raw.length;r++){
    var row=raw[r],fg2=trim(row[FG]);
    if(!fg2||isDec(fg2))continue;
    var fgl=fg2.toLowerCase();
    if(!batchInfo[fgl])batchInfo[fgl]={qty:0,unit:''};
    var bq=pn(row[BQTY]);if(bq>0&&batchInfo[fgl].qty===0)batchInfo[fgl].qty=bq;
    var bu=trim(String(row[BUNT]||''));if(bu&&!batchInfo[fgl].unit)batchInfo[fgl].unit=bu;
  }
  for(var r=hr+1;r<raw.length;r++){
    var row=raw[r],fg2=trim(row[FG]);
    if(!fg2||isDec(fg2))continue;
    var fgl=fg2.toLowerCase();
    if(!map[fgl]){var info=batchInfo[fgl]||{qty:0,unit:''};map[fgl]={name:fg2,qty:info.qty,unit:info.unit,ingredients:[]};}
    var mat=trim(row[MAT]);if(!mat)continue;
    var rate=pn(row[RATE]);
    if(rate<=0&&pm[mat.toLowerCase()])rate=pm[mat.toLowerCase()].price;
    var qty=pn(row[QTY]),unit=trim(row[UNIT]);
    map[fgl].ingredients.push({material:mat,qty:qty,unit:unit,rate:rate,lineTotal:qty*rate});
  }
  return map;
}

function getRecipeMap(){
  try{
    var d=rd('Recipes');
    return buildRecMap(d.raw, rd('RM_Price').rows.map(function(r){
      return{name:trim(r['Material_Name_Eng']||r['Material_Name']||''),price:pn(r['Price_Per_Unit']||0),unit:trim(r['Unit']||'')};
    }));
  }catch(e){return{};}
}

// ══════════════════════════════════════
// ── BUILD SUB-RM MAP ──
// ══════════════════════════════════════
function buildSubRmMap(){
  var map={};
  try{
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sub_RM_Recipes');
    if(!sheet)return{};
    var data=sheet.getDataRange().getValues();
    if(data.length<2)return{};
    var headers=data[0].map(function(h){return h.toString().toLowerCase().trim();});
    var idxName     = headers.indexOf('sub_rm_name');
    var idxIng      = headers.indexOf('ingredient_name');
    var idxQty      = headers.indexOf('qty_required');
    var idxUnit     = headers.indexOf('unit');
    var idxYieldQty = headers.indexOf('batch_output_qty');
    var idxYieldUnit= headers.indexOf('batch_output_unit');
    // fallback column names (space variants)
    if(idxYieldQty  < 0) idxYieldQty  = headers.indexOf('batch output qty');
    if(idxYieldUnit < 0) idxYieldUnit = headers.indexOf('batch output unit');

    for(var i=1;i<data.length;i++){
      var row=data[i];
      var name=row[idxName]?row[idxName].toString().trim():'';
      if(!name)continue;
      if(!map[name]){
        map[name]={
          name      : name,
          outputQty : parseFloat(row[idxYieldQty]) || 1,
          outputUnit: row[idxYieldUnit] ? row[idxYieldUnit].toString().trim() : 'Kg',
          ingredients: []
        };
      }
      if(row[idxIng]){
        map[name].ingredients.push({
          material: row[idxIng].toString().trim(),
          qty     : parseFloat(row[idxQty]) || 0,
          unit    : row[idxUnit] ? row[idxUnit].toString().trim() : ''
        });
      }
    }
  }catch(e){Logger.log('Error in buildSubRmMap: '+e.message);}
  return map;
}

// ══════════════════════════════════════
// ── FG CRUD ──
// ══════════════════════════════════════
function saveFG(fg){
  var name=trim(fg.name);if(!name)throw new Error('Name required');
  ensureSh('FG_List',['FG_Name','Category','GST_Pct','Notes']);
  var d=rd('FG_List'),ni=d.h.indexOf('FG_Name');
  var row=[name,trim(fg.category),pn(fg.gst),trim(fg.notes)];
  for(var i=d.ds;i<d.raw.length;i++){
    if(kw(trim(String(d.raw[i][ni]||'')))===kw(name)){
      d.s.getRange(i+1,1,1,4).setValues([row]);
      return{success:true,message:'"'+name+'" updated.'};
    }
  }
  d.s.appendRow(row);
  return{success:true,message:'"'+name+'" added.'};
}

function deleteFG(name){
  name=trim(name);
  var d=rd('FG_List'),ni=d.h.indexOf('FG_Name');
  for(var i=d.raw.length-1;i>=d.ds;i--){
    if(kw(trim(String(d.raw[i][ni]||'')))===kw(name)){
      d.s.deleteRow(i+1);
      return{success:true,message:'Deleted.'};
    }
  }
  throw new Error('"'+name+'" not found');
}

// ══════════════════════════════════════
// ── RM CRUD ──
// ══════════════════════════════════════
function saveRM(row, mode){
  var eng  = trim(row.name);
  var hin  = trim(row.nameHi || '');
  var cat  = trim(row.category);
  var unit = trim(row.unit);
  var price= pn(row.price);
  if(!eng||!unit) throw new Error('Name and Unit required');
  if(!hin){try{hin=LanguageApp.translate(eng,'en','hi');}catch(e){hin='';}}
  ensureSh('RM_Price',['Material_Name_Eng','Material_Name_Hindi','Category','Unit','Price_Per_Unit']);
  var d  = rd('RM_Price');
  var ni = ci(d.h,['materialnameeng','materialname']);if(ni<0)ni=0;
  var rowData = [eng,hin,cat,unit,price];

  if(mode==='upsert'||mode==='add'){
    // For upsert: update if same name+category exists
    if(mode==='upsert'){
      for(var i=d.ds;i<d.raw.length;i++){
        if(kw(trim(String(d.raw[i][ni]||'')))===kw(eng)){
          d.s.getRange(i+1,1,1,5).setValues([rowData]);
          return{success:true,message:'"'+eng+'" updated.',nameHi:hin};
        }
      }
    }
    // For add: check duplicate
    if(mode==='add'){
      var catIdx=ci(d.h,['category','dept','department']);if(catIdx<0)catIdx=2;
      for(var i=d.ds;i<d.raw.length;i++){
        if(kw(trim(String(d.raw[i][ni]||'')))===kw(eng)&&kw(trim(String(d.raw[i][catIdx]||'')))===kw(cat)){
          throw new Error('"'+eng+'" already exists in '+cat+'! Duplicate entries are not allowed.');
        }
      }
    }
    d.s.appendRow(rowData);
    return{success:true,message:'"'+eng+'" added.',nameHi:hin};
  }
  d.s.appendRow(rowData);
  return{success:true,message:'"'+eng+'" saved.',nameHi:hin};
}

function deleteRM(name){
  name=trim(name);
  var d=rd('RM_Price'),ni=ci(d.h,['materialnameeng','materialname']);if(ni<0)ni=0;
  for(var i=d.raw.length-1;i>=d.ds;i--){
    if(kw(trim(String(d.raw[i][ni]||'')))===kw(name)){
      d.s.deleteRow(i+1);
      return{success:true,message:'Deleted.'};
    }
  }
  throw new Error('"'+name+'" not found');
}

// ══════════════════════════════════════
// ── OVERHEADS CRUD ──
// ══════════════════════════════════════
function saveOverhead(cat,subCat,basis,pct){
  cat=trim(cat);subCat=trim(subCat);basis=trim(basis)||'% of RM Cost';pct=pn(pct);
  if(!cat)throw new Error('Category required');
  ensureSh('Overheads',['Category','Sub_Category','Calculation_Basis','Percentage']);
  var d=rd('Overheads');
  var ci2=d.h.indexOf('Category');if(ci2<0)ci2=0;
  var sci=d.h.indexOf('Sub_Category');
  for(var i=d.ds;i<d.raw.length;i++){
    var rc=kw(trim(String(d.raw[i][ci2]||'')));
    var rs=kw(trim(String(sci>=0?d.raw[i][sci]:'')||''));
    if(rc===kw(cat)&&rs===kw(subCat)){
      d.s.getRange(i+1,1,1,4).setValues([[cat,subCat,basis,pct]]);
      return{success:true,message:'"'+cat+'" updated.',overheads:_readOHRows()};
    }
  }
  d.s.appendRow([cat,subCat,basis,pct]);
  return{success:true,message:'"'+cat+'" added.',overheads:_readOHRows()};
}

function updateOverhead(origCat,origSubCat,nc,nSubCat,basis,pct){
  origCat=trim(origCat);origSubCat=trim(origSubCat);nc=trim(nc);nSubCat=trim(nSubCat);
  basis=trim(basis)||'% of RM Cost';pct=pn(pct);
  if(!nc)throw new Error('Overhead name required');
  var d=rd('Overheads');
  var ci2=d.h.indexOf('Category');if(ci2<0)ci2=0;
  var sci=d.h.indexOf('Sub_Category');
  for(var i=d.ds;i<d.raw.length;i++){
    var rc=kw(trim(String(d.raw[i][ci2]||'')));
    var rs=kw(trim(String(sci>=0?d.raw[i][sci]:'')||''));
    if(rc===kw(origCat)&&rs===kw(origSubCat)){
      d.s.getRange(i+1,1,1,4).setValues([[nc,nSubCat,basis,pct]]);
      return{success:true,message:'"'+nc+'" updated.',overheads:_readOHRows()};
    }
  }
  d.s.appendRow([nc,nSubCat,basis,pct]);
  return{success:true,message:'"'+nc+'" saved.',overheads:_readOHRows()};
}

function deleteOverhead(cat,subCat){
  cat=trim(cat);subCat=trim(subCat);
  var d=rd('Overheads');
  var ci2=d.h.indexOf('Category');if(ci2<0)ci2=0;
  var sci=d.h.indexOf('Sub_Category');
  for(var i=d.raw.length-1;i>=d.ds;i--){
    var rc=kw(trim(String(d.raw[i][ci2]||'')));
    var rs=kw(trim(String(sci>=0?d.raw[i][sci]:'')||''));
    if(rc===kw(cat)&&rs===kw(subCat)){
      d.s.deleteRow(i+1);
      return{success:true,message:'Deleted.',overheads:_readOHRows()};
    }
  }
  throw new Error('"'+cat+'" not found');
}

function getOverheads(){return _readOHRows();}

// ══════════════════════════════════════
// ── DELETE RECIPE ──
// ══════════════════════════════════════
function deleteRecipe(fgName){
  fgName=trim(fgName);
  var d=rd('Recipes'),fi=ci(d.h,['fgname','productname']);if(fi<0)fi=0;
  var n=0;
  for(var i=d.raw.length-1;i>=d.ds;i--){
    if(kw(trim(String(d.raw[i][fi]||'')))===kw(fgName)){d.s.deleteRow(i+1);n++;}
  }
  if(!n)throw new Error('"'+fgName+'" not found');
  return{success:true,message:'Deleted.'};
}

// ══════════════════════════════════════
// ── SUB-RM CRUD ──
// ══════════════════════════════════════
function saveSubRM(subName, ings, replace, batchQty, batchUnit){
  subName   = trim(subName);
  batchQty  = pn(batchQty) || 1;
  batchUnit = trim(batchUnit) || 'Kg';
  if(!subName)       throw new Error('Sub-RM Name is required');
  if(!ings || !ings.length) throw new Error('No ingredients provided for Sub-RM');

  var s=ensureSh('Sub_RM_Recipes',['Sub_RM_Name','Ingredient_Name','Qty_Required','Unit','Batch_Output_Qty','Batch_Output_Unit']);
  var d=rd('Sub_RM_Recipes');
  var ni=d.h.indexOf('Sub_RM_Name');if(ni<0)ni=0;

  if(replace){
    for(var i=d.raw.length-1;i>=d.ds;i--){
      if(kw(trim(String(d.raw[i][ni]||'')))===kw(subName)){s.deleteRow(i+1);}
    }
  }

  ings.forEach(function(g,idx){
    var m=trim(g.material),u=trim(g.unit)||'Kg',q=pn(g.qty);
    if(!m)return;
    s.appendRow([subName,m,q,u,idx===0?batchQty:'',idx===0?batchUnit:'']);
  });

  // Calculate Normalized Cost and update RM_Price
  var normalizedCpu=0;
  try{
    var rmRows=rd('RM_Price').rows;
    var pm={};
    rmRows.forEach(function(r){
      var nm=trim(r['Material_Name_Eng']||r['Material_Name']||'');
      if(nm)pm[kw(nm)]=pn(r['Price_Per_Unit']||r['Price']||0);
    });
    var totalBatchCost=0;
    ings.forEach(function(g){
      totalBatchCost+=pn(g.qty)*(pm[kw(trim(g.material))]||0);
    });
    normalizedCpu=batchQty>0?totalBatchCost/batchQty:totalBatchCost;

    var rmSh=sh('RM_Price'),rd2=rd('RM_Price');
    var ni2=ci(rd2.h,['materialnameeng','materialname']);if(ni2<0)ni2=0;
    var found=false;
    for(var j=rd2.ds;j<rd2.raw.length;j++){
      if(kw(trim(String(rd2.raw[j][ni2]||'')))===kw(subName)){
        rmSh.getRange(j+1,5).setValue(normalizedCpu);
        rmSh.getRange(j+1,4).setValue(batchUnit);
        found=true;break;
      }
    }
    if(!found){rmSh.appendRow([subName,'','Sub-RM',batchUnit,normalizedCpu]);}
  }catch(e){Logger.log('Error updating RM_Price for Sub-RM: '+e);}

  return{success:true,message:'"'+subName+'" Sub-RM saved. Cost/Unit: ₹'+normalizedCpu.toFixed(2),cpu:normalizedCpu};
}

function deleteSubRM(name){
  name=trim(name);
  var d=rd('Sub_RM_Recipes'),ni=d.h.indexOf('Sub_RM_Name');if(ni<0)ni=0;
  var n=0;
  for(var i=d.raw.length-1;i>=d.ds;i--){
    if(kw(trim(String(d.raw[i][ni]||'')))===kw(name)){d.s.deleteRow(i+1);n++;}
  }
  if(!n)throw new Error('"'+name+'" not found');
  return{success:true,message:'Deleted.'};
}

// ══════════════════════════════════════
// ── CALCULATE COST ──
// ══════════════════════════════════════
function calculateCost(fgName, sellingPrice, ohOvr){
  sellingPrice=pn(sellingPrice);
  ohOvr=ohOvr||{};

  var rmRows=rd('RM_Price').rows;
  var ohRows=_readOHRows();
  var rmList=rmRows.map(function(r){
    return{
      name : trim(r['Material_Name_Eng']||r['Material_Name']||''),
      price: pn(r['Price_Per_Unit']||r['Price']||0),
      unit : trim(r['Unit']||'')
    };
  });

  var recMap=buildRecMap(rd('Recipes').raw,rmList);
  var subMap=buildSubRmMap();

  // Build Price Map (include Sub-RM computed costs)
  var pm={};
  rmList.forEach(function(r){if(r.name)pm[r.name.toLowerCase()]={price:r.price,unit:r.unit};});
  Object.keys(subMap).forEach(function(sn){
    var totalCost=0;
    subMap[sn].ingredients.forEach(function(g){
      var k=g.material.toLowerCase();
      if(pm[k])totalCost+=pn(g.qty)*pm[k].price;
    });
    var outQty=pn(subMap[sn].outputQty)||1;
    pm[sn.toLowerCase()]={price:totalCost/outQty,unit:subMap[sn].outputUnit||'Kg'};
  });

  var fgKey=kw(fgName);
  var rk=Object.keys(recMap).find(function(k){return kw(k)===fgKey;});
  if(!rk)throw new Error('No recipe for "'+fgName+'". Add recipe first.');
  var recipe=recMap[rk];

  var fgMeta={gst:0,category:''};
  try{
    var fl=rd('FG_List').rows.find(function(r){return kw(trim(r['FG_Name']))===fgKey;});
    if(fl){fgMeta.gst=pn(fl['GST_Pct']);fgMeta.category=trim(fl['Category']);}
  }catch(e){}

  // Component A: Recipe RM Cost
  var lines=[],batchRMTotal=0,miss=[];
  recipe.ingredients.forEach(function(g){
    var k=g.material.toLowerCase();
    var entry=pm[k];
    var rate=entry?entry.price:(pn(g.rate)>0?pn(g.rate):0);
    var line=pn(g.qty)*rate;
    if(rate<=0)miss.push(g.material);
    batchRMTotal+=line;
    lines.push({
      material:g.material,qty:pn(g.qty),
      unit:entry?entry.unit:(g.unit||''),
      rate:rate,line:line,missing:rate<=0,
      isSubRM:!!(subMap[g.material])
    });
  });

  // Normalize RM to Per Unit
  var bq=pn(recipe.qty)||1;
  var cpRMPerUnit=batchRMTotal/bq;

  // Component B: Overheads (applied Per Unit)
  var ohL=[],ohTotalPerUnit=0;
  ohRows.forEach(function(o){
    var cat=trim(o.category);
    var subCat=trim(o.subCategory||'');
    // Filter: skip if sub-category set and doesn't match FG category
    if(subCat&&kw(subCat)!==kw(fgMeta.category))return;

    var overrideKey=ckGS(cat)+(subCat?'_'+ckGS(subCat):'');
    var pct=(ohOvr[overrideKey]!==undefined&&ohOvr[overrideKey]!=='')?pn(ohOvr[overrideKey]):o.percentage;
    var basis=trim(o.basis)||'% of RM Cost';
    var baseValue=basis.toLowerCase().indexOf('selling')>-1?sellingPrice:cpRMPerUnit;
    var amt=baseValue*pct/100;
    ohTotalPerUnit+=amt;
    ohL.push({category:cat,subCategory:subCat,basis:basis,pct:pct,base:baseValue,amount:amt});
  });

  var subTotalPerUnit=cpRMPerUnit+ohTotalPerUnit;
  var gstAmt=subTotalPerUnit*fgMeta.gst/100;
  var profitPU=sellingPrice>0?sellingPrice-subTotalPerUnit:0;
  var netPct=sellingPrice>0?(profitPU/sellingPrice)*100:0;

  return{
    fgName:fgName, fgCategory:fgMeta.category,
    batchQty:bq, batchUnit:recipe.unit||'Unit',
    lineItems:lines, missing:miss,
    batchRM:batchRMTotal, cpRMPerUnit:cpRMPerUnit,
    ohLines:ohL, ohTotal:ohTotalPerUnit,
    finCostPerUnit:subTotalPerUnit,
    gstPct:fgMeta.gst, gstAmt:gstAmt,
    totalWithGst:subTotalPerUnit+gstAmt,
    costPerUnit:subTotalPerUnit,
    sellingPrice:sellingPrice,
    profitPerUnit:profitPU, netProfitPct:netPct,
    timestamp:new Date().toLocaleString('en-IN',{timeZone:'Asia/Kolkata'})
  };
}

// ══════════════════════════════════════
// ── SAVE CALCULATION LOG ──
// ══════════════════════════════════════
function saveCalculation(r){
  try{writeREADME();}catch(e){Logger.log('README:'+e);}
  var s=ensureSh('Cost_Log',[
    'Timestamp','Product','Category','Batch Qty','Unit',
    'Batch RM','Overhead','GST','Total Cost','Cost/Unit',
    'Selling/Unit','Profit/Unit','Margin%','Net Profit%'
  ]);
  var marginPct=r.marginPct||r.netProfitPct||0;
  s.appendRow([
    r.timestamp, r.fgName, r.fgCategory||'', r.batchQty, r.batchUnit,
    +r.batchRM.toFixed(2), +r.ohTotal.toFixed(2), +r.gstAmt.toFixed(2),
    +r.totalWithGst.toFixed(2), +r.costPerUnit.toFixed(2),
    r.sellingPrice, +(r.profitPerUnit||0).toFixed(2),
    marginPct.toFixed(1)+'%', (r.netProfitPct||0).toFixed(1)+'%'
  ]);
  return true;
}

// ══════════════════════════════════════
// ── README SHEET ──
// ══════════════════════════════════════
function writeREADME(){
  var s=ensureSh('README',['What We Calculate','How It Works (Plain English)','Example / Notes']);
  s.clearContents();
  s.appendRow(['What We Calculate','How It Works (Plain English)','Example / Notes']);
  s.getRange(1,1,1,3).setFontWeight('bold').setBackground('#3B1F0E').setFontColor('#FFF');
  var rows=[
    ['About this file','MithaiCalc updates this sheet every time you save a calculation.','Last updated: '+new Date().toLocaleString('en-IN',{timeZone:'Asia/Kolkata'})],
    ['','',''],
    ['--- OVERHEADS ---','',''],
    ['Sub_Category blank','Overhead applies to ALL products.','e.g. Labour 10% applies to every product'],
    ['Sub_Category = Methai','Overhead applies only to Methai products.','e.g. Methai Packaging applies only when FG Category = Methai'],
    ['','',''],
    ['--- COST FORMULA ---','',''],
    ['RM Cost per Unit','Total batch RM ÷ batch qty','e.g. ₹66.98 ÷ 4 Pcs = ₹16.74/Pc'],
    ['Overhead per Unit','Sum of all overheads applied to RM/Selling base','e.g. Labour 10% of ₹16.74 = ₹1.674'],
    ['Final Costing (A+B)','RM per unit + Overhead per unit','e.g. ₹16.74 + ₹1.68 = ₹18.42/Pc'],
    ['GST','Final cost per unit × GST%',''],
    ['Total Cost per Unit','Final cost per unit + GST',''],
    ['','',''],
    ['--- SAVE BUTTONS ---','',''],
    ['💾 Save Recipe','Saves ingredients to Recipes sheet only','Use for normal finished goods'],
    ['💾 Save Sub Raw Material Recipe','Saves to Sub_RM_Recipes + Recipes + FG_List + RM_Price','Use for composite ingredients like Masala Mix'],
  ];
  rows.forEach(function(r){s.appendRow(r);});
  s.autoResizeColumns(1,3);
}

// ══════════════════════════════════════
// ── SYNC ALL RECIPE RATES ──
// ══════════════════════════════════════
function updateAllRecipeRates(){
  try{
    var subMap=buildSubRmMap();
    var rmRows=rd('RM_Price').rows;
    var pm={};
    rmRows.forEach(function(r){
      var nm=trim(r['Material_Name_Eng']||r['Material_Name']||'');
      if(nm)pm[kw(nm)]=pn(r['Price_Per_Unit']||r['Price']||0);
    });
    var rmSh=sh('RM_Price'),rd2=rd('RM_Price');
    var ni2=ci(rd2.h,['materialnameeng','materialname']);if(ni2<0)ni2=0;
    var updated=0;
    Object.keys(subMap).forEach(function(sn){
      var s=subMap[sn],totalCost=0,oq=pn(s.outputQty)||1;
      s.ingredients.forEach(function(g){totalCost+=pn(g.qty)*(pm[kw(trim(g.material))]||0);});
      var cpu=totalCost/oq;
      for(var j=rd2.ds;j<rd2.raw.length;j++){
        if(kw(trim(String(rd2.raw[j][ni2]||'')))===kw(sn)){
          rmSh.getRange(j+1,5).setValue(cpu);updated++;break;
        }
      }
    });
    return{success:true,message:'Rates synced! '+updated+' Sub-RM prices updated.'};
  }catch(e){
    throw new Error('Sync failed: '+e.message);
  }
}

// ══════════════════════════════════════
// ── BATCH OVERWRITE HELPER ──
// ══════════════════════════════════════
function batchOverwriteRows(sheetName, idColIndex, idValue, newRowsData){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName(sheetName);
  if(!sheet)return;
  var fullData=sheet.getDataRange().getValues();
  var headers=fullData[0];
  var filteredData=fullData.filter(function(row,index){
    if(index===0)return true;
    return kw(String(row[idColIndex]||''))!==kw(String(idValue||''));
  });
  newRowsData.forEach(function(row){filteredData.push(row);});
  sheet.clearContents();
  if(filteredData.length>0&&headers.length>0){
    sheet.getRange(1,1,filteredData.length,headers.length).setValues(filteredData);
  }
}