/* Ledger — spelling (Tools ▸ Spelling, F7) and AutoCorrect for cell entries.
 * Words are checked against Hunspell's English word list (SCOWL-derived, dict/en_US.words);
 * suggestions come from edit distance ranked by how common a word is.
 */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, O = L.ops, F = L.formula;
  const { h } = L;
  const ui = L.ui;
  const SP = (L.spell = {});
  const opt = (k, def) => { const o = L.app && L.app.opts; return o && o[k] !== undefined ? o[k] !== false : def !== false; };
  SP.refreshSoon = () => {};

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
  SP.reloadCustom = () => { SP._custom = null; };
  const ignoredAll = new Set();
  const cache = new Map();
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
  const restoreCase = (sugg, orig) => (isUpper(orig) && orig.length > 1 ? sugg.toUpperCase() : /^\p{Lu}/u.test(orig) && !/^\p{Lu}/u.test(sugg) ? sugg[0].toUpperCase() + sugg.slice(1) : sugg);
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

  /* ================= AutoCorrect ================= */
  const AC = (L.autocorrect = {});
  const DEFAULTS = `(c)=©|(r)=®|(tm)=™|...=…|:)=☺|:-)=☺|:(=☹|:-(=☹|:|=😐|-->=→|<--=←|abbout=about|abotu=about|abouta=about a|aboutit=about it|abscence=absence|accesories=accessories|accidant=accident|accomodate=accommodate|accross=across|acheive=achieve|acheived=achieved|acheiving=achieving|acn=can|acommodate=accommodate|acomodate=accommodate|actualyl=actually|additinal=additional|addtional=additional|adequit=adequate|adn=and|advanage=advantage|affraid=afraid|afterthe=after the|againstt he=against the|aganist=against|aggresive=aggressive|agian=again|agreemeent=agreement|agreemnet=agreement|agressive=aggressive|ahppen=happen|ahve=have|allwasy=always|allwyas=always|almots=almost|almsot=almost|alomst=almost|alot=a lot|alraedy=already|alreayd=already|alreday=already|alwasy=always|alwats=always|alway=always|alwyas=always|amde=made|Ameria=America|amke=make|amkes=makes|anbd=and|andone=and one|andteh=and the|andthe=and the|anothe=another|anual=annual|apparant=apparent|apparrent=apparent|appearence=appearance|appeares=appears|applicaiton=application|applicaitons=applications|applyed=applied|appointiment=appointment|approrpiate=appropriate|approrpriate=appropriate|aquisition=acquisition|aquisitions=acquisitions|arguement=argument|arguements=arguments|arn't=aren't|arond=around|artical=article|articel=article|asdvertising=advertising|assistent=assistant|asthe=as the|atention=attention|atmospher=atmosphere|attentioin=attention|atthe=at the|audeince=audience|audiance=audience|availalbe=available|awya=away|aywa=away|bakc=back|balence=balance|ballance=balance|baout=about|bcak=back|beacuse=because|becasue=because|becaus=because|becausea=because a|becauseof=because of|becausethe=because the|becauseyou=because you|becomeing=becoming|becomming=becoming|becuase=because|becuse=because|befoer=before|beggining=beginning|begining=beginning|beginnig=beginning|beleive=believe|belive=believe|beteen=between|beween=between|bewteen=between|bilsters=blisters|bve=be|boxs=boxes|brodcast=broadcast|butthe=but the|bve=be|cafe=café|caharcter=character|calcullated=calculated|calulated=calculated|candidtae=candidate|candidtaes=candidates|catagory=category|categiory=category|certian=certain|challange=challenge|challanges=challenges|chaneg=change|chanegs=changes|changable=changeable|charachter=character|charachters=characters|charactor=character|charecter=character|charector=character|cheif=chief|chekc=check|chnage=change|cieling=ceiling|circut=circuit|claer=clear|claered=cleared|claerly=clearly|cliant=client|cliche=cliché|cna=can|colection=collection|comanies=companies|comany=company|comapnies=companies|comapny=company|combintation=combination|comited=committed|comittee=committee|commadn=command|comming=coming|commitee=committee|committe=committee|committment=commitment|committments=commitments|committy=committee|comntain=contain|comntains=contains|compair=compare|compleated=completed|compleatly=completely|compleatness=completeness|completly=completely|completness=completeness|composate=composite|comtain=contain|comtains=contains|comunicate=communicate|comunity=community|condolances=condolences|conected=connected|conferance=conference|confirmmation=confirmation|considerit=considerate|considerite=considerate|consonent=consonant|conspiricy=conspiracy|consultent=consultant|convertable=convertible|cooperatoin=cooperation|corproation=corporation|corproations=corporations|corruptable=corruptible|cotten=cotton|coudl=could|could of been=could have been|could of had=could have had|couldnt=couldn't|couldthe=could the|cpoy=copy|creme=crème|ctaegory=category|cusotmer=customer|cusotmers=customers|cutsomer=customer|cutsomers=customers|cxan=can|danceing=dancing|dcument=document|deatils=details|decison=decision|decisons=decisions|decor=décor|defendent=defendant|definately=definitely|deptartment=department|desicion=decision|desicions=decisions|desision=decision|desisions=decisions|developement=development|developped=developed|develpment=development|devleop=develop|devrom=devform|didint=didn't|didnot=did not|didnt=didn't|difefrent=different|diferences=differences|differance=difference|differances=differences|differant=different|differemt=different|differnt=different|diffrent=different|directer=director|directers=directors|directiosn=direction|disatisfied=dissatisfied|discoverd=discovered|disign=design|dispaly=display|dissonent=dissonant|distribusion=distribution|divsion=division|docuement=documents|docuemnt=document|documetn=document|documnet=document|documnets=documents|doese=does|doesnt=doesn't|doign=doing|doimg=doing|doind=doing|dollers=dollars|donig=doing|dont=don't|dosn't=doesn't|driveing=driving|drnik=drink|eclair=éclair|efel=feel|effecient=efficient|efort=effort|eforts=efforts|ehr=her|eligable=eligible|embarass=embarrass|emigre=émigré|encouraing=encouraging|enought=enough|equippment=equipment|equivalant=equivalent|esle=else|especally=especially|especialyl=especially|espesially=especially|excellant=excellent|excercise=exercise|exchagne=exchange|exchagnes=exchanges|excitment=excitement|exhcange=exchange|exhcanges=exchanges|experiance=experience|experienc=experience|exprience=experience|exprienced=experienced|eyt=yet|facade=façade|faeture=feature|faetures=features|familair=familiar|familar=familiar|familliar=familiar|fammiliar=familiar|feild=field|feilds=fields|fianlly=finally|fidn=find|finalyl=finally|firends=friends|firts=first|follwo=follow|follwoing=following|fora=for a|foriegn=foreign|forthe=for the|forwrd=forward|forwrds=forwards|foudn=found|foward=forward|fowards=forwards|freind=friend|frmo=from|fromthe=from the|furneral=funeral|fwe=few|garantee=guarantee|gaurd=guard|gemeral=general|gerat=great|geting=getting|gettin=getting|gievn=given|giveing=giving|gloabl=global|goign=going|gonig=going|govenment=government|goverment=government|gruop=group|gruops=groups|grwo=grow|guidlines=guidelines|hadbeen=had been|haev=have|hapen=happen|hapened=happened|hapening=happening|hapens=happens|happend=happened|hasbeen=has been|hasnt=hasn't|havebeen=have been|haveing=having|hda=had|hearign=hearing|hed=he'd|helpfull=helpful|herat=heart|hesaid=he said|hewas=he was|hge=he|hismelf=himself|hlep=help|hsa=has|hsi=his|hte=the|htere=there|htese=these|htey=they|hting=thing|htink=think|htis=this|hvae=have|hvaing=having|hwich=which|i=I|i'd=I'd|i'll=I'll|i'm=I'm|i've=I've|idae=idea|idaes=ideas|identofy=identify|imagenary=imaginary|imagin=imagine|immediatly=immediately|immediatley=immediately|imporatnt=important|importamt=important|importent=important|importnat=important|impossable=impossible|improvemnt=improvement|improvment=improvement|includ=include|indecate=indicate|indenpendence=independence|indenpendent=independent|indepedent=independent|independance=independence|independant=independent|influance=influence|infomation=information|informatoin=information|inital=initial|instaleld=installed|insted=instead|insurence=insurance|inteh=in the|interum=interim|inthe=in the|inwhich=in which|isthe=is the|itis=it is|ititial=initial|itnerest=interest|itnerested=interested|itneresting=interesting|itnerests=interests|itwas=it was|iwll=will|iwth=with|jsut=just|knowldge=knowledge|knowlege=knowledge|knwo=know|knwon=known|knwos=knows|konw=know|konwn=known|konws=knows|labratory=laboratory|lastyear=last year|learnign=learning|lenght=length|levle=level|libary=library|librarry=library|librery=library|liek=like|liekd=liked|lieutenent=lieutenant|liev=live|likly=likely|lisense=license|littel=little|litttle=little|liuke=like|liveing=living|loev=love|lonly=lonely|lookign=looking|maintenence=maintenance|makeing=making|managment=management|mantain=maintain|marraige=marriage|memeber=member|merchent=merchant|mesage=message|mesages=messages|mkae=make|mkaes=makes|mkaing=making|moeny=money|morgage=mortgage|mroe=more|mysefl=myself|myu=my|naive=naïve|necassarily=necessarily|necassary=necessary|neccessarily=necessarily|neccessary=necessary|necesarily=necessarily|necesary=necessary|negotiaing=negotiating|nkow=know|nothign=nothing|nver=never|nwe=new|nwo=now|obvioulsy=obviously|ocasion=occasion|ocassion=occasion|occurence=occurrence|occurrance=occurrence|ocur=occur|oeprator=operator|ofits=of its|ofthe=of the|oging=going|ohter=other|omre=more|oneof=one of|onepoint=one point|onthe=on the|onyl=only|oppasite=opposite|opperation=operation|oppertunity=opportunity|opposate=opposite|opposible=opposable|opposit=opposite|oppotunities=opportunities|oppotunity=opportunity|orginization=organization|orginized=organized|otehr=other|otu=out|outof=out of|overthe=over the|owrk=work|owuld=would|oxident=oxidant|papaer=paper|parliment=parliament|partof=part of|paymetn=payment|paymetns=payments|pciture=picture|peice=piece|peices=pieces|peolpe=people|peopel=people|percentof=percent of|percentto=percent to|performence=performance|perhasp=perhaps|perhpas=perhaps|permanant=permanent|perminent=permanent|personalyl=personally|pleasent=pleasant|poeple=people|porblem=problem|porblems=problems|porvide=provide|possable=possible|postition=position|potentialy=potentially|pregnent=pregnant|presance=presence|probelm=problem|probelms=problems|prominant=prominent|protege=protégé|protoge=protégé|psoition=position|ptogress=progress|puting=putting|pwoer=power|quater=quarter|quaters=quarters|quesion=question|quesions=questions|questioms=questions|questiosn=questions|questoin=question|quetion=question|quetions=questions|realyl=really|reccomend=recommend|reccommend=recommend|receieve=receive|recieve=receive|recieved=received|recieving=receiving|recomend=recommend|recomendation=recommendation|recomendations=recommendations|recomended=recommended|reconize=recognize|recrod=record|religous=religious|reluctent=reluctant|remeber=remember|reommend=recommend|representatiive=representative|representives=representatives|represetned=represented|represnt=represent|reserach=research|resollution=resolution|resorces=resources|respomd=respond|respomse=response|responce=response|responsability=responsibility|responsable=responsible|responsibile=responsible|responsiblity=responsibility|restaraunt=restaurant|restuarant=restaurant|reult=result|reveiw=review|reveiwing=reviewing|rumers=rumors|rwite=write|rythm=rhythm|saidhe=said he|saidit=said it|saidthat=said that|saidthe=said the|scedule=schedule|sceduled=scheduled|seance=séance|secratary=secretary|sectino=section|seh=she|selectoin=selection|sentance=sentence|separeate=separate|seperate=separate|sercumstances=circumstances|shcool=school|shesaid=she said|shineing=shining|shiped=shipped|shoudl=should|shoudln't=shouldn't|should of been=should have been|should of had=should have had|shouldnt=shouldn't|showinf=showing|signifacnt=significant|simalar=similar|similiar=similar|simpyl=simply|sincerly=sincerely|sitll=still|smae=same|smoe=some|soem=some|sohw=show|soical=social|somethign=something|someting=something|somewaht=somewhat|somthing=something|somtimes=sometimes|soudn=sound|soudns=sounds|speach=speech|specificaly=specifically|specificalyl=specifically|statment=statement|statments=statements|stnad=stand|stopry=story|stoyr=story|stpo=stop|strentgh=strength|stroy=story|struggel=struggle|strugle=struggle|studnet=student|successfull=successful|successfuly=successfully|successfulyl=successfully|sucess=success|sucessfull=successful|sufficiant=sufficient|suposed=supposed|suppose to=supposed to|supposingly=supposedly|suprise=surprise|suprised=surprised|swiming=swimming|tahn=than|taht=that|talekd=talked|talkign=talking|tath=that|tecnical=technical|teh=the|tehy=they|tellt he=tell the|termoil=turmoil|tghe=the|tghis=this|thansk=thanks|thats=that's|thatthe=that the|thecompany=the company|thefirst=the first|thegovernment=the government|themself=themselves|themselfs=themselves|thenew=the new|theri=their|thesame=the same|thetwo=the two|thgat=that|thge=the|thier=their|thigsn=things|thisyear=this year|thme=them|thna=than|thne=then|thnig=thing|thnigs=things|threatend=threatened|thsi=this|thsoe=those|thta=that|tihs=this|timne=time|tiogether=together|tje=the|tjhe=the|tkae=take|tkaes=takes|tkaing=taking|tlaking=talking|todya=today|togehter=together|tomorow=tomorrow|tongiht=tonight|tonihgt=tonight|totaly=totally|totalyl=totally|tothe=to the|towrad=toward|traditionalyl=traditionally|transfered=transferred|truely=truly|truley=truly|tryed=tried|tthe=the|tyhat=that|tyhe=the|udnerstand=understand|understnad=understand|undert he=under the|unitedstates=United States|unliek=unlike|unpleasently=unpleasantly|untill=until|untilll=until|useing=using|usualyl=usually|veyr=very|virtualyl=virtually|vrey=very|vulnerible=vulnerable|waht=what|warrent=warrant|wasnt=wasn't|watn=want|wehn=when|werre=were|whcih=which|wherre=where|whic=which|whihc=which|whta=what|wief=wife|wierd=weird|wihch=which|wiht=with|windoes=windows|withe=with|wiull=will|wnat=want|wnated=wanted|wnats=wants|woh=who|wohle=whole|wokr=work|woudl=would|woudln't=wouldn't|would of been=would have been|would of had=would have had|wouldbe=would be|wouldnt=wouldn't|wriet=write|writting=writing|wrod=word|wroet=wrote|wroking=working|wtih=with|wuould=would|wya=way|yera=year|yeras=years|yersa=years|yoiu=you|youare=you are|youve=you've|ytou=you|yuo=you|yuor=your`;
  const EXCEPTIONS = ['a.', 'abbr.', 'abs.', 'acad.', 'acct.', 'adj.', 'adv.', 'al.', 'apr.', 'approx.', 'assn.', 'aug.', 'ave.', 'b.a.', 'c.', 'ca.', 'cf.', 'co.', 'corp.', 'dec.', 'dept.', 'dr.', 'e.g.', 'ed.', 'esp.', 'est.', 'etc.', 'ex.', 'feb.', 'fig.', 'figs.', 'fri.', 'gen.', 'i.e.', 'inc.', 'jan.', 'jr.', 'jul.', 'jun.', 'ltd.', 'mar.', 'max.', 'min.', 'misc.', 'mon.', 'mr.', 'mrs.', 'ms.', 'mt.', 'no.', 'nov.', 'oct.', 'p.', 'pp.', 'prof.', 'ref.', 'rev.', 'sat.', 'sep.', 'sept.', 'sr.', 'st.', 'sun.', 'tel.', 'thu.', 'tue.', 'u.s.', 'vol.', 'vs.', 'wed.'];
  AC.list = function () {
    const base = new Map();
    for (const pair of DEFAULTS.split('|')) { const i = pair.indexOf('='); if (i > 0) base.set(pair.slice(0, i), pair.slice(i + 1)); }
    base.delete('i'); base.delete(':|'); base.delete(':)'); base.delete(':-)'); base.delete(':('); base.delete(':-(');
    const user = L.store.get('acList', null);
    if (user) { for (const k of user.del || []) base.delete(k); for (const [k, v] of user.add || []) base.set(k, v); }
    return base;
  };
  let LIST = null;
  AC.reload = () => { LIST = AC.list(); };
  AC.setEntry = function (from, to) { const u = L.store.get('acList', { add: [], del: [] }); u.add = (u.add || []).filter((x) => x[0] !== from); u.del = (u.del || []).filter((x) => x !== from); if (to != null) u.add.push([from, to]); else u.del.push(from); L.store.set('acList', u); AC.reload(); };
  AC.exceptions = () => L.store.get('acExceptions', EXCEPTIONS);
  const DAYS = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
  /**
   * AutoCorrect a typed text entry (not formulas): replacement list, TWo INitial CApitals,
   * capitalised day names and first letters of sentences, accidental cAPS LOCK.
   */
  AC.applyCell = function (text) {
    if (!text || /^[=+\-@']/.test(text) || !opt('acReplace', true) && !opt('acTwoCaps', true)) return text;
    if (!LIST) AC.reload();
    let out = text;
    if (opt('acReplace', true)) {
      out = out.replace(/(^|[\s(])([^\s()]+?)(?=$|[\s,.;:!?)])/g, (m, pre, w) => {
        if (LIST.has(w)) return pre + LIST.get(w);
        const lw = w.toLowerCase();
        if (LIST.has(lw) && /^[a-z]/.test(LIST.get(lw))) { const r = LIST.get(lw); return pre + (w[0] === w[0].toUpperCase() && w[0] !== w[0].toLowerCase() ? r[0].toUpperCase() + r.slice(1) : r); }
        return m;
      });
      for (const k of ['(c)', '(r)', '(tm)', '...', '-->', '<--', '==>', '<==', '<=>']) if (out.indexOf(k) >= 0 && LIST.has(k)) out = out.split(k).join(LIST.get(k));
    }
    if (opt('acTwoCaps', true)) out = out.replace(/\b([A-Z])([A-Z])([a-z]{2,})\b/g, (m, a, b, rest) => a + b.toLowerCase() + rest);
    if (opt('acDays', true)) out = out.replace(/\b(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/g, (m) => m[0].toUpperCase() + m.slice(1));
    if (opt('acCapsLock', true)) out = out.replace(/\b([a-z])([A-Z]{2,})\b/g, (m, a, rest) => a.toUpperCase() + rest.toLowerCase());
    if (opt('acSentence', false)) out = out.replace(/(^|[.!?]\s+)([a-z])/g, (m, pre, ch) => pre + ch.toUpperCase());
    void DAYS;
    return out;
  };

  /* ================= Spelling dialog (cells and comments) ================= */
  const WORD = /[\p{L}][\p{L}\p{M}'’\-]*[\p{L}\p{M}]|[\p{L}]/gu;
  /** misspellings in a text: [{word, index}] */
  function findErrors(text, list) {
    const out = [];
    WORD.lastIndex = 0;
    let m;
    while ((m = WORD.exec(text))) {
      const w = m[0];
      if (/^(https?|www|mailto)/i.test(w) || /\.\w+\//.test(text.slice(m.index - 1, m.index + w.length + 2))) continue;
      if (opt('ignoreInternet', true) && /[@\/\\]/.test(text.slice(Math.max(0, m.index - 1), m.index + w.length + 1))) continue;
      if (SP.ignoredAll.has(w.toLowerCase())) continue;
      if (!SP.okWord(w, list)) out.push({ word: w, index: m.index });
    }
    return out;
  }
  SP.ignoredAll = new Set();
  SP.findErrors = findErrors;
  SP.check = async function () {
    const G = L.grid, A = L.app;
    const sh = G.sheet(), wb = G.wb;
    let list = (A.opts && A.opts.dictLang) === 'en_GB' ? 'en_GB' : 'en_US';
    ui.busy(true, 'Loading the dictionary...');
    try { await LEX.load(list); } catch (e) { ui.busy(false); ui.msg('The spelling dictionary could not be loaded (' + (e.message || e) + ').', { icon: 'warn' }); return; }
    ui.busy(false);
    /* cells to check: the selection when it is more than one cell, otherwise the sheet from the active cell, wrapping */
    const rg0 = G.range();
    const single = G.ranges().length === 1 && rg0.r1 === rg0.r2 && rg0.c1 === rg0.c2;
    const s0 = G.sel();
    const cells = [];
    sh.rows.forEach((row, r) => { if (!row) return; row.cells.forEach((cell, c) => { if (!cell || typeof cell.v !== 'string' || cell.f != null || cell.am) return; if (!single && !G.ranges().some((x) => M.rangeContains(x, r, c))) return; cells.push({ r, c }); }); });
    if (single) { const k = cells.findIndex((x) => x.r > s0.r || (x.r === s0.r && x.c >= s0.c)); if (k > 0) cells.push(...cells.splice(0, k)); }
    let ci = 0, hits = [], hi = 0, changes = 0;
    const changeAll = new Map();
    const next = () => {
      for (;;) {
        if (hi < hits.length) return true;
        if (ci >= cells.length) return false;
        const it = cells[ci++];
        const v = sh.val(it.r, it.c);
        if (typeof v !== 'string') continue;
        hits = findErrors(v, list).map((x) => Object.assign(x, it));
        hi = 0;
        /* Change All replacements apply silently */
        hits = hits.filter((x) => { if (changeAll.has(x.word)) { replaceIn(x, changeAll.get(x.word)); return false; } return true; });
      }
    };
    const replaceIn = (x, to) => {
      const v = sh.val(x.r, x.c);
      if (typeof v !== 'string') return;
      const re = new RegExp('(^|[^\\p{L}])' + x.word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?=$|[^\\p{L}])', 'u');
      const nv = v.replace(re, (m, pre) => pre + to);
      if (nv !== v) { O.tx(wb, 'Spelling', () => O.setValue(sh, x.r, x.c, nv)); changes++; }
    };
    if (!next()) { ui.msg(changes ? 'The spelling check is complete for the entire sheet.' : 'The spelling check is complete for the entire sheet.', { icon: 'info' }); return; }
    const word = h('input', { type: 'text', style: 'width:100%' });
    const ctx = h('div', { class: 'sp-ctx' });
    const sugg = h('select', { size: '6', style: 'width:100%' });
    const btn = (label, fn) => ui.button(label, fn, { style: 'width:100%' });
    let cur = null;
    const show = () => {
      cur = hits[hi];
      G.select(cur.r, cur.c);
      word.value = cur.word;
      const v = sh.val(cur.r, cur.c);
      ctx.textContent = '';
      ctx.append(document.createTextNode(v.slice(Math.max(0, cur.index - 40), cur.index)), h('b', { style: 'color:#c00', text: cur.word }), document.createTextNode(v.slice(cur.index + cur.word.length, cur.index + cur.word.length + 40)));
      sugg.textContent = '';
      const s = SP.suggest(cur.word, list, 8);
      for (const x of s.length ? s : ['(No Suggestions)']) sugg.appendChild(h('option', { value: x, text: x }));
      sugg.selectedIndex = 0;
      sugg.disabled = !s.length;
    };
    const advance = () => { hi++; if (!next()) { d.close(null); ui.msg('The spelling check is complete for the entire sheet.', { icon: 'info' }); G.paint(); return; } show(); };
    const pick = () => (word.value !== cur.word ? word.value : sugg.disabled ? null : sugg.value);
    const LANGS = [['en_US', 'English (U.S.)'], ['en_GB', 'English (U.K.)']];
    const langName = (k) => (LANGS.find((x) => x[0] === k) || LANGS[0])[1];
    const langSel = ui.select(LANGS, list, async (v) => {
      /* switch dictionaries: load the list, remember the choice, re-check the current cell */
      ui.busy(true, 'Loading the dictionary...');
      try { await LEX.load(v); } catch (e) { ui.busy(false); ui.msg('The spelling dictionary could not be loaded (' + (e.message || e) + ').', { icon: 'warn' }); langSel.value = list; return; }
      ui.busy(false);
      list = v;
      if (A.opts) { A.opts.dictLang = v; if (A.saveOpts) A.saveOpts(); }
      d.el.querySelector('.dlg-ttl').textContent = 'Spelling: ' + langName(v);
      hits = findErrors(sh.val(cur.r, cur.c), list).map((x) => Object.assign(x, { r: cur.r, c: cur.c }));
      hi = 0;
      if (!hits.length) { hi = 1; advance(); } else show();
    });
    const d = ui.dialog({
      title: 'Spelling: ' + langName(list), width: 520,
      body: h('div', { class: 'sp-grid' },
        h('div', { class: 'col' }, h('label', { text: 'Not in Dictionary:' }), ctx, word, h('label', { text: 'Suggestions:' }), sugg, ui.field('Dictionary &language:', langSel)),
        h('div', { class: 'col sp-btns' },
          btn('&Ignore Once', advance), btn('I&gnore All', () => { SP.ignoredAll.add(cur.word.toLowerCase()); hits = hits.filter((x, k) => k <= hi || x.word.toLowerCase() !== cur.word.toLowerCase()); advance(); }),
          btn('&Add to Dictionary', () => { SP.addWord(cur.word); advance(); }),
          btn('&Change', () => { const v = pick(); if (v != null) { replaceIn(cur, v); hits = findErrors(sh.val(cur.r, cur.c), list).map((x) => Object.assign(x, { r: cur.r, c: cur.c })).filter((x) => x.index > cur.index + v.length - cur.word.length - 1); hi = -1; } advance(); }),
          btn('Change A&ll', () => { const v = pick(); if (v != null) { changeAll.set(cur.word, v); replaceIn(cur, v); } advance(); }),
          btn('Auto&Correct', () => { const v = pick(); if (v == null) return; L.autocorrect.setEntry(cur.word, v); replaceIn(cur, v); advance(); }))),
      buttons: [{ label: '&Options...', onClick: () => { L.dlg.options('Spelling'); return false; } }, { label: 'Cancel' }],
    });
    sugg.addEventListener('dblclick', () => { const v = pick(); if (v != null) { replaceIn(cur, v); } advance(); });
    show();
  };
  SP.addWord = function (w) { const list = L.store.get('customDict', []) || []; if (!list.includes(w)) list.push(w); L.store.set('customDict', list); SP.reloadCustom(); };
})(typeof window !== 'undefined' ? window : globalThis);
