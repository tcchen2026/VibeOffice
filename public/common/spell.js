/* VibeOffice — the spelling engine shared by Quire and Ledger: the offline word lists (common/dict/,
 * Hunspell/SCOWL-derived, ranked by frequency), checking, suggestions by edit distance, the custom dictionary
 * and Ignore All, and AutoCorrect's replacement list. Each app has its own spelling UI on top (Quire: underlines, grammar and its dialog;
 * Ledger: Tools ▸ Spelling); an app that draws underlines sets SP.refreshSoon to redraw them.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const SP = (L.spell = L.spell || {});
  const opt = (k, def) => { const o = L.app && L.app.opts; return o && o[k] !== undefined ? o[k] !== false : def !== false; };
  SP.refreshSoon = SP.refreshSoon || (() => {});

  /* ================= lexicon ================= */
  const LEX = (L.lex = { lists: {}, loading: {}, failed: {} });
  const LEVEL = '0123456789';
  /** load a front-coded list: "<prefix-length char><rest><level digit>" per line */
  LEX.load = function (name) {
    if (LEX.lists[name]) return Promise.resolve(LEX.lists[name]);
    if (LEX.loading[name]) return LEX.loading[name];
    LEX.loading[name] = fetch('../common/dict/' + name + '.words').then((r) => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.text(); }).then((txt) => {
      const map = new Map();
      let prev = '';
      const P = '0123456789abcdefghijklmnopqrstuvwxyz';
      for (const ln of txt.split('\n')) {
        if (!ln) continue;
        const k = P.indexOf(ln[0]);
        const w = prev.slice(0, k) + ln.slice(1, -1);
        map.set(w, LEVEL.indexOf(ln[ln.length - 1]));
        prev = w;
      }
      /* lower-case index so "Paris" at sentence start and "PARIS" both resolve */
      const lower = new Map();
      for (const [w, lv] of map) { const l = w.toLowerCase(); if (!lower.has(l) || lower.get(l).lv > lv) lower.set(l, { w, lv }); }
      LEX.lists[name] = { map, lower };
      L.bus.emit('lexicon-loaded', name);
      return LEX.lists[name];
    }).catch((e) => { LEX.failed[name] = String(e && e.message || e); LEX.loading[name] = null; throw e; });
    return LEX.loading[name];
  };
  /** which word list a language tag uses (null: no English proofing tools for it) */
  LEX.listFor = function (lang) {
    lang = String(lang || 'en-US').toLowerCase();
    if (!/^en\b/.test(lang)) return null;
    return /^en-(gb|au|nz|ie|za|in|sg|hk|jm|bz|tt|zw|ph|my)/.test(lang) ? 'en_GB' : 'en_US';
  };
  LEX.ready = (name) => !!LEX.lists[name];

  /* ================= checking words ================= */
  const custom = () => { if (!SP._custom) SP._custom = new Set((L.store.get('customDict', []) || []).map((x) => String(x).toLowerCase())); return SP._custom; };
  const ignoredAll = new Set();
  const cache = new Map();
  /** after the custom dictionary changed elsewhere (Tools ▸ Options): forget the cached answers */
  SP.reloadCustom = () => { SP._custom = null; cache.clear(); SP.refreshSoon(); };
  /** Add to Dictionary: the word is spelled correctly from now on, in every app (a setting per app, L.store) */
  SP.addWord = function (w) {
    const list = L.store.get('customDict', []) || [];
    if (!list.some((x) => String(x).toLowerCase() === String(w).toLowerCase())) list.push(String(w));
    L.store.set('customDict', list);
    SP.reloadCustom();
  };
  /** Ignore All: accept the word for the rest of the session */
  SP.ignoreAll = (w) => { ignoredAll.add(String(w).toLowerCase()); cache.clear(); };
  const isUpper = (w) => w === w.toUpperCase() && w !== w.toLowerCase();
  const isCap = (w) => w[0] === w[0].toUpperCase() && w.slice(1) === w.slice(1).toLowerCase();
  /** true when the word is spelled correctly (or should not be checked) */
  SP.okWord = function (word, list) {
    const lex = LEX.lists[list];
    if (!lex) return true;
    let w = word.replace(/’/g, "'");
    if (w.length < 2) return true;
    if (/\d/.test(w)) return opt('ignoreNum', true) || /^\d/.test(w);
    if (isUpper(w) && opt('ignoreUpper', true)) return true;
    const key = list + '\u0000' + w;
    if (cache.has(key)) return cache.get(key);
    const r = checkWord(w, lex);
    cache.set(key, r);
    return r;
  };
  function checkWord(w, lex) {
    const lw = w.toLowerCase();
    if (custom().has(lw) || ignoredAll.has(lw)) return true;
    if (lex.map.has(w)) return true;
    /* sentence-initial capitals and ALL CAPS accept the lower-case entry; a capitalised entry ("Paris") needs its capital */
    if ((isCap(w) || isUpper(w)) && lex.map.has(lw)) return true;
    if (isUpper(w)) { const e = lex.lower.get(lw); if (e) return true; }
    if (isCap(w)) { const e = lex.lower.get(lw); if (e && isCap(e.w)) return true; }
    /* possessive of a known word (James's, CEO's) */
    const m = /^(.+?)'s?$/i.exec(w);
    if (m && m[1].length > 1 && (lex.map.has(m[1]) || lex.lower.has(m[1].toLowerCase()))) return true;
    /* hyphenated compounds are fine when every part is */
    if (w.includes('-')) return w.split('-').every((part) => !part || part.length < 2 || checkWord(part, lex));
    /* accented spellings of listed words: café, naïve, résumé, façade */
    const plain = w.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (plain !== w) return checkWord(plain, lex);
    /* AutoCorrect replacements the user taught are accepted words too */
    return false;
  }

  /* ================= suggestions ================= */
  const ALPHA = "abcdefghijklmnopqrstuvwxyz'";
  /* Derived from Peter Norvig's spelling corrector spell.py (norvig.com/spell.py), edits1 (MIT License, Copyright
     (c) 2007-2016 Peter Norvig; common/licenses/norvig-spell-LICENSE.txt), in JavaScript */
  function edits1(w) {
    const out = new Set();
    for (let i = 0; i <= w.length; i++) {
      const a = w.slice(0, i), b = w.slice(i);
      if (b) out.add(a + b.slice(1));
      if (b.length > 1) out.add(a + b[1] + b[0] + b.slice(2));
      for (const c of ALPHA) { if (b) out.add(a + c + b.slice(1)); out.add(a + c + b); }
    }
    out.delete(w);
    return out;
  }
  const KEYS = ['qwertyuiop', 'asdfghjkl', 'zxcvbnm'];
  const keyPos = {};
  KEYS.forEach((row, r) => { for (let c = 0; c < row.length; c++) keyPos[row[c]] = [r, c + r * 0.5]; });
  const near = (a, b) => { const p = keyPos[a], q = keyPos[b]; return p && q && Math.abs(p[0] - q[0]) <= 1 && Math.abs(p[1] - q[1]) <= 1.5; };
  /** Damerau-Levenshtein distance with cheaper adjacent-key and doubled-letter mistakes */
  function dist(a, b) {
    const n = a.length, m = b.length;
    const d = [];
    for (let i = 0; i <= n; i++) { d[i] = [i]; }
    for (let j = 0; j <= m; j++) d[0][j] = j;
    for (let i = 1; i <= n; i++) for (let j = 1; j <= m; j++) {
      const sub = a[i - 1] === b[j - 1] ? 0 : near(a[i - 1], b[j - 1]) ? 0.8 : 1;
      let v = Math.min(d[i - 1][j] + (i > 1 && a[i - 1] === a[i - 2] ? 0.6 : 1), d[i][j - 1] + (j > 1 && b[j - 1] === b[j - 2] ? 0.6 : 1), d[i - 1][j - 1] + sub);
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) v = Math.min(v, d[i - 2][j - 2] + 0.7);
      d[i][j] = v;
    }
    return d[n][m];
  }
  /** sugg in orig's case: all capitals, or a capital first letter */
  const restoreCase = (SP.restoreCase = (sugg, orig) => (isUpper(orig) && orig.length > 1 ? sugg.toUpperCase() : /^\p{Lu}/u.test(orig) && !/^\p{Lu}/u.test(sugg) ? sugg[0].toUpperCase() + sugg.slice(1) : sugg));
  SP.suggest = function (word, list, max) {
    max = max || 8;
    const lex = LEX.lists[list];
    if (!lex) return [];
    const w = word.replace(/’/g, "'");
    const lw = w.toLowerCase();
    const found = new Map();
    const consider = (cand, base) => {
      const e = lex.lower.get(cand);
      if (!e) return;
      const real = e.w;
      if (real === w) return;
      const score = dist(lw, cand) * 10 + e.lv * 1.6 + base + (real.length < 3 ? 4 : 0);
      if (!found.has(real) || found.get(real) > score) found.set(real, score);
    };
    /* AutoCorrect knows many typical slips */
    const ac = L.autocorrect ? L.autocorrect.list() : null;
    if (ac && ac.has(lw)) { const r = ac.get(lw); if (!/\s/.test(r)) found.set(r, -100); else found.set(r, -50); }
    const e1 = edits1(lw);
    for (const c of e1) consider(c, 0);
    if (found.size < 6 && lw.length <= 14) {
      for (const c of e1) { if (lex.lower.has(c)) continue; for (const c2 of edits1(c)) if (c2.length > 1) consider(c2, 6); }
    }
    /* run-together words: "alot" → "a lot", "infact" → "in fact" */
    for (let i = 1; i < lw.length; i++) {
      const a = lw.slice(0, i), b = lw.slice(i);
      if ((a.length > 1 || a === 'a' || a === 'i') && b.length > 1 && lex.lower.has(a) && lex.lower.has(b) && lex.lower.get(a).lv <= 4 && lex.lower.get(b).lv <= 5) {
        const score = 11 + (lex.lower.get(a).lv + lex.lower.get(b).lv) * 1.2;
        const s = (a === 'i' ? 'I' : lex.lower.get(a).w) + ' ' + lex.lower.get(b).w;
        if (!found.has(s) || found.get(s) > score) found.set(s, score);
      }
    }
    return Array.from(found.entries()).sort((x, y) => x[1] - y[1]).slice(0, max).map(([s]) => (/^\p{Lu}/u.test(s) && !isUpper(w) ? s : restoreCase(s, w)));
  };
  /* ================= AutoCorrect's default list (Quire and Ledger) ================= */
  /* Derived from Apache OpenOffice extras/source/autotext/lang/en-US/acor_en-US.dat (DocumentList.xml; Apache License
     2.0): its entries less those our AutoFormat options cover (--, 1/2, 1/4, 3/4) and "I.e.", with "cutsomers" made
     "customers" (it gave "customer"). entry=replacement, separated by | */
  const ACOR_AOO = `-->=→|->=→|...=…|(C)=©|(R)=®|(tm)=™|<-=←|<--=←|<-->=↔|<->=↔|abotu=about|abouta=about a|aboutit=about it|abscence=absence|accesories=accessories|accidant=accident|accomodate=accommodate|accordingto=according to|accross=across|acheive=achieve|acheived=achieved|acheiving=achieving|acn=can|acommodate=accommodate|acomodate=accommodate|actualyl=actually|additinal=additional|addtional=additional|adequit=adequate|adequite=adequate|adn=and|advanage=advantage|affraid=afraid|afterthe=after the|againstt he=against the|aggresive=aggressive|agian=again|agreemeent=agreement|agreemeents=agreements|agreemnet=agreement|agreemnets=agreements|agressive=aggressive|ahppen=happen|ahve=have|allready=already|allwasy=always|allwyas=always|almots=almost|almsot=almost|alomst=almost|alot=a lot|alraedy=already|alreayd=already|alreday=already|alwasy=always|alwats=always|alway=always|alwyas=always|amde=made|Ameria=America|amke=make|amkes=makes|anbd=and|andone=and one|andt he=and the|andteh=and the|andthe=and the|anothe=another|anual=annual|apparant=apparent|apparrent=apparent|appearence=appearance|appeares=appears|applicaiton=application|applicaitons=applications|applyed=applied|appointiment=appointment|approrpiate=appropriate|approrpriate=appropriate|aquisition=acquisition|aquisitions=acquisitions|arguement=argument|arguements=arguments|arond=around|artical=article|articel=article|asdvertising=advertising|askt he=ask the|atention=attention|atmospher=atmosphere|attentioin=attention|atthe=at the|audeince=audience|audiance=audience|availalbe=available|awya=away|aywa=away|bakc=back|balence=balance|ballance=balance|baout=about|bcak=back|beacuse=because|becasue=because|becaus=because|becausea=because a|becauseof=because of|becausethe=because the|becauseyou=because you|becomeing=becoming|becomming=becoming|becuase=because|becuse=because|befoer=before|beggining=beginning|begining=beginning|beginining=beginning|beleiev=believe|beleieve=believe|beleif=belief|beleive=believe|beleived=believed|beleives=believes|benifit=benefit|benifits=benefits|betwen=between|beutiful=beautiful|blase=blasé|boxs=boxes|brodcast=broadcast|butthe=but the|bve=be|byt he=by the|caharcter=character|calcullated=calculated|calulated=calculated|candidtae=candidate|candidtaes=candidates|catagory=category|categiory=category|certian=certain|challange=challenge|challanges=challenges|chaneg=change|chanegs=changes|changable=changeable|changeing=changing|changng=changing|charachter=character|charachters=characters|charactor=character|charecter=character|charector=character|cheif=chief|chekc=check|chnage=change|cieling=ceiling|circut=circuit|claer=clear|claered=cleared|claerly=clearly|cliant=client|cliche=cliché|cna=can|colection=collection|comanies=companies|comany=company|comapnies=companies|comapny=company|combintation=combination|comited=committed|comittee=committee|commadn=command|comming=coming|commitee=committee|committe=committee|committment=commitment|committments=commitments|committy=committee|comntain=contain|comntains=contains|compair=compare|compleated=completed|compleatly=completely|compleatness=completeness|completly=completely|completness=completeness|composate=composite|comtain=contain|comtains=contains|comunicate=communicate|comunity=community|condolances=condolences|conected=connected|conferance=conference|confirmmation=confirmation|considerit=considerate|considerite=considerate|consonent=consonant|conspiricy=conspiracy|consultent=consultant|convertable=convertible|cooparate=cooperate|cooporate=cooperate|corproation=corporation|corproations=corporations|corruptable=corruptible|cotten=cotton|coudl=could|could of been=could have been|could of had=could have had|couldthe=could the|cpoy=copy|creme=crème|ctaegory=category|cusotmer=customer|cusotmers=customers|cutsomer=customer|cutsomers=customers|cxan=can|danceing=dancing|dcument=document|deatils=details|decison=decision|decisons=decisions|decor=décor|defendent=defendant|definately=definitely|deptartment=department|desicion=decision|desicions=decisions|desision=decision|desisions=decisions|develeoprs=developers|devellop=develop|develloped=developed|develloper=developer|devellopers=developers|develloping=developing|devellopment=development|devellopments=developments|devellops=develop|develope=develop|developement=development|developements=developments|developor=developer|developors=developers|develpment=development|diaplay=display|didnot=did not|difefrent=different|diferences=differences|differance=difference|differances=differences|differant=different|differemt=different|differnt=different|diffrent=different|directer=director|directers=directors|directiosn=direction|disatisfied=dissatisfied|discoverd=discovered|disign=design|dispaly=display|dissonent=dissonant|distribusion=distribution|divsion=division|docuement=document|docuemnt=document|documetn=document|documnet=document|documnets=documents|doe snot=does not|doese=does|doign=doing|doimg=doing|doind=doing|dollers=dollars|donig=doing|driveing=driving|drnik=drink|efel=feel|effecient=efficient|efort=effort|eforts=efforts|ehr=her|eligable=eligible|embarass=embarrass|enought=enough|equippment=equipment|equivalant=equivalent|esle=else|especally=especially|especialyl=especially|espesially=especially|excellant=excellent|excercise=exercise|exchagne=exchange|exchagnes=exchanges|excitment=excitement|exhcange=exchange|exhcanges=exchanges|experiance=experience|experienc=experience|exprience=experience|exprienced=experienced|eyt=yet|faeture=feature|faetures=features|familair=familiar|familar=familiar|familliar=familiar|fammiliar=familiar|feild=field|feilds=fields|fianlly=finally|fidn=find|finalyl=finally|firends=friends|firts=first|follwo=follow|follwoing=following|fora=for a|foriegn=foreign|forthe=for the|forwrd=forward|forwrds=forwards|foudn=found|foward=forward|fowards=forwards|freind=friend|freindly=friendly|freinds=friends|frmo=from|fromt he=from the|fromthe=from the|furneral=funeral|fwe=few|garantee=guarantee|gaurd=guard|gemeral=general|gerat=great|geting=getting|gettin=getting|gievn=given|giveing=giving|gloabl=global|goign=going|gonig=going|govenment=government|goverment=government|gruop=group|gruops=groups|grwo=grow|guidlines=guidelines|hadbeen=had been|haev=have|hapen=happen|hapened=happened|hapening=happening|hapens=happens|happend=happened|hasbeen=has been|havebeen=have been|haveing=having|hearign=hearing|helpfull=helpful|herat=heart|hesaid=he said|hewas=he was|hge=he|hismelf=himself|hlep=help|hsa=has|hsi=his|hte=the|htere=there|htese=these|htey=they|hting=thing|htink=think|htis=this|hvae=have|hvaing=having|hwich=which|i=I|i snot=is not|idae=idea|idaes=ideas|identofy=identify|ihs=his|imediate=immediate|imediatly=immediately|immediatly=immediately|importent=important|importnat=important|impossable=impossible|improvemnt=improvement|improvment=improvement|includ=include|indecate=indicate|indenpendence=independence|indenpendent=independent|indepedent=independent|independance=independence|independant=independent|influance=influence|infomation=information|informatoin=information|inital=initial|instaleld=installed|insted=instead|insurence=insurance|int he=in the|inteh=in the|interum=interim|inthe=in the|inwhich=in which|isthe=is the|itis=it is|ititial=initial|itnerest=interest|itnerested=interested|itneresting=interesting|itnerests=interests|iwll=will|iwth=with|jsut=just|jugment=judgment|knowldge=knowledge|knowlege=knowledge|knwo=know|knwon=known|knwos=knows|konwn=known|konws=knows|labratory=laboratory|lastyear=last year|learnign=learning|lenght=length|levle=level|libary=library|librarry=library|librery=library|liek=like|liekd=liked|lieutenent=lieutenant|liev=live|likly=likely|lisense=license|littel=little|litttle=little|liuke=like|liveing=living|loev=love|lonly=lonely|lookign=looking|maintenence=maintenance|makeing=making|managment=management|mantain=maintain|marraige=marriage|may of been=may have been|may of had=may have had|memeber=member|merchent=merchant|mesage=message|mesages=messages|might of been=might have been|might of had=might have had|mispell=misspell|mispelling=misspelling|mispellings=misspellings|mkae=make|mkaes=makes|mkaing=making|moeny=money|morgage=mortgage|mroe=more|must of been=must have been|must of had=must have had|mysefl=myself|myu=my|necassarily=necessarily|necassary=necessary|neccessarily=necessarily|neccessary=necessary|necesarily=necessarily|necesary=necessary|negotiaing=negotiating|nkow=know|nothign=nothing|nver=never|nwe=new|nwo=now|obediant=obedient|ocasion=occasion|occassion=occasion|occured=occurred|occurence=occurrence|occurrance=occurrence|ocur=occur|odf=ODF|oeprator=operator|ofits=of its|oft he=of the|ofthe=of the|oging=going|ohter=other|omre=more|oneof=one of|onepoint=one point|ont he=on the|onthe=on the|onyl=only|oppasite=opposite|opperation=operation|oppertunity=opportunity|opposate=opposite|opposible=opposable|opposit=opposite|oppotunities=opportunities|oppotunity=opportunity|orginization=organization|orginized=organized|otehr=other|otu=out|outof=out of|overthe=over the|owrk=work|owuld=would|oxident=oxidant|papaer=paper|parliment=parliament|partof=part of|paymetn=payment|paymetns=payments|pciture=picture|peice=piece|peices=pieces|peolpe=people|peopel=people|percentof=percent of|percentto=percent to|performence=performance|perhasp=perhaps|perhpas=perhaps|permanant=permanent|perminent=permanent|personalyl=personally|pleasent=pleasant|poeple=people|porblem=problem|porblems=problems|porvide=provide|possable=possible|postition=position|potentialy=potentially|pregnent=pregnant|presance=presence|probelm=problem|probelms=problems|prominant=prominent|psoition=position|ptogress=progress|puting=putting|pwoer=power|quater=quarter|quaters=quarters|quesion=question|quesions=questions|questioms=questions|questiosn=questions|questoin=question|quetion=question|quetions=questions|realyl=really|reccomend=recommend|reccommend=recommend|receieve=receive|recieve=receive|recieved=received|recieving=receiving|recomend=recommend|recomendation=recommendation|recomendations=recommendations|recomended=recommended|reconize=recognize|recrod=record|religous=religious|reluctent=reluctant|remeber=remember|reommend=recommend|representativs=representatives|representives=representatives|represetned=represented|represnt=represent|reserach=research|resollution=resolution|resorces=resources|respomd=respond|respomse=response|responce=response|responsability=responsibility|responsable=responsible|responsibile=responsible|responsiblity=responsibility|restaraunt=restaurant|restuarant=restaurant|reult=result|reveiw=review|reveiwing=reviewing|rumers=rumors|rwite=write|rythm=rhythm|saidhe=said he|saidit=said it|saidt he=said the|saidthat=said that|saidthe=said the|scedule=schedule|sceduled=scheduled|seance=séance|secratary=secretary|sectino=section|seh=she|selectoin=selection|sentance=sentence|separeate=separate|sercumstances=circumstances|shcool=school|shesaid=she said|shineing=shining|shiped=shipped|shoudl=should|should of been=should have been|should of had=should have had|showinf=showing|signifacnt=significant|simalar=similar|similiar=similar|simpyl=simply|sincerly=sincerely|sitll=still|smae=same|smoe=some|soem=some|sohw=show|soical=social|somethign=something|someting=something|somewaht=somewhat|somthing=something|somtimes=sometimes|sot hat=so that|soudn=sound|soudns=sounds|speach=speech|specificaly=specifically|specificalyl=specifically|statment=statement|statments=statements|stnad=stand|stopry=story|stoyr=story|stpo=stop|strentgh=strength|stroy=story|struggel=struggle|strugle=struggle|studnet=student|successfull=successful|successfuly=successfully|successfulyl=successfully|sucess=success|sucessfull=successful|sufficiant=sufficient|suposed=supposed|suppossed=supposed|suprised=surprised|swiming=swimming|tahn=than|taht=that|talekd=talked|talkign=talking|tath=that|tecnical=technical|teh=the|tehy=they|tellt he=tell the|termoil=turmoil|tets=test|tghe=the|tghis=this|thansk=thanks|thatt he=that the|thatthe=that the|thecompany=the company|thefirst=the first|thegovernment=the government|their are=there are|their is=there is|themself=themselves|themselfs=themselves|thenew=the new|theri=their|thesame=the same|thetwo=the two|thgat=that|thge=the|thier=their|thigsn=things|thisyear=this year|thme=them|thna=than|thne=then|thnig=thing|thnigs=things|threatend=threatened|thsi=this|thsoe=those|thta=that|tihs=this|timne=time|tiogether=together|tje=the|tjhe=the|tkae=take|tkaes=takes|tkaing=taking|tlaking=talking|todya=today|togehter=together|toldt he=told the|tomorow=tomorrow|tongiht=tonight|tonihgt=tonight|totaly=totally|totalyl=totally|tothe=to the|towrad=toward|traditionalyl=traditionally|transfered=transferred|truely=truly|truley=truly|tryed=tried|tthe=the|tyhat=that|tyhe=the|udnerstand=understand|understnad=understand|undert he=under the|UnitedStates=United States|unliek=unlike|unpleasently=unpleasantly|untill=until|untilll=until|useing=using|usualyl=usually|veyr=very|virtualyl=virtually|vrey=very|vulnerible=vulnerable|wa snot=was not|waht=what|warrent=warrant|watn=want|wehn=when|werre=were|whcih=which|whent he=when the|wherre=where|whic=which|whicht he=which the|whihc=which|who;ve=who’ve|whta=what|wief=wife|wierd=weird|wihch=which|wiht=with|will of been=will have been|willbe=will be|windoes=windows|witha=with a|withe=with|witht he=with the|withthe=with the|wiull=will|wnat=want|wnated=wanted|wnats=wants|woh=who|wohle=whole|wokr=work|woudl=would|would of been=would have been|would of had=would have had|wouldbe=would be|wriet=write|writting=writing|wrod=word|wroet=wrote|wroking=working|wtih=with|wuould=would|wya=way|yera=year|yeras=years|yersa=years|yoiu=you|youare=you are|ytou=you|yuo=you|yuor=your`;
  /* our own: symbols; the contractions typed without their apostrophe (left out where that spelling is a word of
     its own: cant, wont, well, were, hell, shell, shed, wed, its, lets, ill, id); I in them, and alone; accents in
     loanwords whose plain spelling isn't an English word */
  const ACOR_SYMBOLS = [['(c)', '©'], ['(r)', '®'], ['(tm)', '™'], ['...', '…'], [':)', '☺'], [':-)', '☺'], [':(', '☹'], [':-(', '☹'],
    ['-->', '→'], ['<--', '←'], ['==>', '⇒'], ['<==', '⇐'], ['<=>', '⇔']];
  const CONTRACTIONS = ("aren't can't couldn't didn't doesn't don't hadn't hasn't haven't isn't mightn't mustn't needn't shan't " +
    "shouldn't wasn't weren't won't wouldn't he'd he'll he's it'd it'll it's let's she'd she'll she's that's that'll there's there'll " +
    "they'd they'll they're they've we'd we'll we're we've what's what'll where's who'd who'll who's who've you'd you'll you're you've " +
    "here's how's when's why's").split(' ');
  const PLAIN_WORDS = new Set(['cant', 'wont', 'well', 'were', 'hell', 'shell', 'shed', 'wed', 'its', 'lets', 'ill', 'id', 'hows', 'whens', 'whys', 'heres']);
  const LOANWORDS = [['cafe', 'café'], ['cafes', 'cafés'], ['cliche', 'cliché'], ['cliches', 'clichés'], ['creme', 'crème'], ['decor', 'décor'],
    ['eclair', 'éclair'], ['eclairs', 'éclairs'], ['emigre', 'émigré'], ['emigres', 'émigrés'], ['facade', 'façade'], ['facades', 'façades'],
    ['naive', 'naïve'], ['naivete', 'naïveté'], ['protege', 'protégé'], ['proteges', 'protégés'], ['seance', 'séance'], ['fiance', 'fiancé'],
    ['fiancee', 'fiancée'], ['entree', 'entrée'], ['entrees', 'entrées'], ['saute', 'sauté'], ['sauteed', 'sautéed'], ['touche', 'touché'],
    ['deja vu', 'déjà vu'], ['pinata', 'piñata'], ['jalapeno', 'jalapeño'], ['jalapenos', 'jalapeños'], ['vis-a-vis', 'vis-à-vis'],
    ['a la carte', 'à la carte'], ['creme brulee', 'crème brûlée'], ['tete-a-tete', 'tête-à-tête'], ['doppelganger', 'doppelgänger']];
  /** the default AutoCorrect entries, entry → replacement (each app removes what it doesn't want and adds the user's) */
  SP.autocorrectDefaults = function () {
    const list = new Map();
    for (const pair of ACOR_AOO.split('|')) { const i = pair.indexOf('='); if (i > 0) list.set(pair.slice(0, i), pair.slice(i + 1)); }
    for (const [k, v] of ACOR_SYMBOLS) list.set(k, v);
    for (const c of CONTRACTIONS) { const plain = c.replace("'", ''); if (!PLAIN_WORDS.has(plain) && !list.has(plain)) list.set(plain, c); }
    for (const c of ["i'm", "i'd", "i'll", "i've"]) list.set(c, 'I' + c.slice(1));
    list.set('i', 'I');
    for (const [k, v] of LOANWORDS) if (!list.has(k)) list.set(k, v);
    return list;
  };
  /** AutoCorrect's replacement list: the defaults less the entries an app leaves out (drop), with the user's additions
   *  and deletions (the app's setting 'acList': { add: [[from, to]], del: [from] }) */
  SP.autocorrectList = function (drop) {
    const list = SP.autocorrectDefaults();
    for (const k of drop || []) list.delete(k);
    const user = L.store.get('acList', null);
    if (user) { for (const k of user.del || []) list.delete(k); for (const [k, v] of user.add || []) list.set(k, v); }
    return list;
  };
  /** an entry added or changed (to), or deleted (to == null) */
  SP.setAutocorrectEntry = function (from, to) {
    const u = L.store.get('acList', { add: [], del: [] });
    u.add = (u.add || []).filter((x) => x[0] !== from);
    u.del = (u.del || []).filter((x) => x !== from);
    if (to != null) u.add.push([from, to]); else if (SP.autocorrectDefaults().has(from)) u.del.push(from);
    L.store.set('acList', u);
  };
})(typeof window !== 'undefined' ? window : globalThis);
